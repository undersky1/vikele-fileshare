import {
  BadRequestException,
  Injectable,
  InternalServerErrorException,
  Logger,
  NotFoundException,
} from "@nestjs/common";
import { PrismaService } from "src/prisma/prisma.service";
import { ConfigService } from "src/config/config.service";
import * as crypto from "crypto";
import * as mime from "mime-types";
import { File } from "./file.service";
import { Readable } from "stream";
import { validate as isValidUUID } from "uuid";
import * as archiver from "archiver";

interface SharePointToken {
  access_token: string;
  expires_on: number;
}

@Injectable()
export class SharePointFileService {
  private readonly logger = new Logger(SharePointFileService.name);
  private cachedToken: SharePointToken | null = null;

  // Track in-progress chunked uploads: fileId -> accumulated Buffer chunks
  private chunkBuffers: Record<string, Buffer[]> = {};

  constructor(
    private prisma: PrismaService,
    private config: ConfigService,
  ) {}

  /**
   * Acquire an OAuth2 client_credentials token for Microsoft Graph.
   */
  private async getAccessToken(): Promise<string> {
    // Return cached token if still valid (with 5min buffer)
    if (
      this.cachedToken &&
      Date.now() / 1000 < this.cachedToken.expires_on - 300
    ) {
      return this.cachedToken.access_token;
    }

    const tenantId = this.config.get("sharepoint.tenantId");
    const clientId = this.config.get("sharepoint.clientId");
    const clientSecret = this.config.get("sharepoint.clientSecret");

    const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

    const body = new URLSearchParams({
      client_id: clientId,
      client_secret: clientSecret,
      scope: "https://graph.microsoft.com/.default",
      grant_type: "client_credentials",
    });

    const response = await fetch(tokenUrl, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: body.toString(),
    });

    if (!response.ok) {
      const errorText = await response.text();
      this.logger.error(`SharePoint token error: ${errorText}`);
      throw new InternalServerErrorException(
        "Failed to authenticate with SharePoint",
      );
    }

    const data = (await response.json()) as {
      access_token: string;
      expires_in?: number;
    };
    this.cachedToken = {
      access_token: data.access_token,
      expires_on: data.expires_in
        ? Date.now() / 1000 + data.expires_in
        : Date.now() / 1000 + 3600,
    };

    return this.cachedToken.access_token;
  }

  /**
   * Build the Graph API base URL for the configured drive.
   */
  private getDriveBaseUrl(): string {
    const siteId = this.config.get("sharepoint.siteId");
    const driveId = this.config.get("sharepoint.driveId");

    if (driveId) {
      return `https://graph.microsoft.com/v1.0/drives/${driveId}`;
    }
    return `https://graph.microsoft.com/v1.0/sites/${siteId}/drive`;
  }

  /**
   * Get the folder path prefix inside the drive.
   */
  private getFolderPath(): string {
    const folderPath = this.config.get("sharepoint.folderPath");
    if (!folderPath) return "";
    const normalized = `${folderPath}`.replace(/^\/+|\/+$/g, "");
    return normalized ? `${normalized}/` : "";
  }

  /**
   * Upload a file (or chunk) to SharePoint via Microsoft Graph.
   * Buffers chunks in memory and uploads the complete file when the last chunk arrives.
   */
  async create(
    data: string,
    chunk: { index: number; total: number },
    file: { id?: string; name: string },
    shareId: string,
  ) {
    if (!file.id) {
      file.id = crypto.randomUUID();
    } else if (!isValidUUID(file.id)) {
      throw new BadRequestException("Invalid file ID format");
    }

    const buffer = Buffer.from(data, "base64");

    // Accumulate chunks
    if (chunk.index === 0) {
      this.chunkBuffers[file.id] = [buffer];
    } else {
      if (!this.chunkBuffers[file.id]) {
        throw new InternalServerErrorException(
          "SharePoint upload session not found for this file.",
        );
      }
      this.chunkBuffers[file.id].push(buffer);
    }

    const isLastChunk = chunk.index === chunk.total - 1;

    if (isLastChunk) {
      const fullBuffer = Buffer.concat(this.chunkBuffers[file.id]);
      delete this.chunkBuffers[file.id];

      const remotePath = `${this.getFolderPath()}${shareId}/${file.name}`;
      const token = await this.getAccessToken();
      const uploadUrl = `${this.getDriveBaseUrl()}/root:/${remotePath}:/content`;

      const response = await fetch(uploadUrl, {
        method: "PUT",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/octet-stream",
        },
        body: fullBuffer,
      });

      if (!response.ok) {
        const errorText = await response.text();
        this.logger.error(`SharePoint upload error: ${errorText}`);
        throw new InternalServerErrorException(
          "Failed to upload file to SharePoint",
        );
      }

      const fileSize = fullBuffer.length;

      await this.prisma.file.create({
        data: {
          id: file.id,
          name: file.name,
          size: fileSize.toString(),
          share: { connect: { id: shareId } },
        },
      });
    }

    return file;
  }

  async get(shareId: string, fileId: string): Promise<File> {
    const fileMetaData = await this.prisma.file.findUnique({
      where: { id: fileId },
    });

    if (!fileMetaData) throw new NotFoundException("File not found");

    const remotePath = `${this.getFolderPath()}${shareId}/${fileMetaData.name}`;
    const token = await this.getAccessToken();
    const downloadUrl = `${this.getDriveBaseUrl()}/root:/${remotePath}:/content`;

    const response = await fetch(downloadUrl, {
      method: "GET",
      headers: {
        Authorization: `Bearer ${token}`,
      },
    });

    if (!response.ok) {
      throw new NotFoundException("File not found in SharePoint");
    }

    // Convert web ReadableStream to Node.js Readable
    const nodeStream = Readable.fromWeb(response.body as any);

    return {
      metaData: {
        id: fileId,
        size: fileMetaData.size,
        name: fileMetaData.name,
        shareId: shareId,
        createdAt: fileMetaData.createdAt,
        mimeType:
          mime.contentType(fileMetaData.name.split(".").pop()) ||
          "application/octet-stream",
      },
      file: nodeStream,
    } as File;
  }

  async remove(shareId: string, fileId: string) {
    const fileMetaData = await this.prisma.file.findUnique({
      where: { id: fileId },
    });

    if (!fileMetaData) throw new NotFoundException("File not found");

    const remotePath = `${this.getFolderPath()}${shareId}/${fileMetaData.name}`;
    const token = await this.getAccessToken();
    const deleteUrl = `${this.getDriveBaseUrl()}/root:/${remotePath}`;

    try {
      const response = await fetch(deleteUrl, {
        method: "DELETE",
        headers: {
          Authorization: `Bearer ${token}`,
        },
      });

      if (!response.ok && response.status !== 404) {
        throw new Error("Delete failed");
      }
    } catch (error) {
      this.logger.error("Failed to delete file from SharePoint", error);
    }

    await this.prisma.file.delete({ where: { id: fileId } });
  }

  async deleteAllFiles(shareId: string) {
    const folderPath = `${this.getFolderPath()}${shareId}`;
    const token = await this.getAccessToken();
    const deleteUrl = `${this.getDriveBaseUrl()}/root:/${folderPath}`;

    try {
      await fetch(deleteUrl, {
        method: "DELETE",
        headers: {
          Authorization: `Bearer ${token}`,
        },
      });
    } catch (error) {
      this.logger.error(
        "Failed to delete share folder from SharePoint",
        error,
      );
    }
  }

  async getZip(shareId: string): Promise<Readable> {
    return new Promise<Readable>(async (resolve, reject) => {
      const compressionLevel = this.config.get("share.zipCompressionLevel");

      try {
        const files = await this.prisma.file.findMany({
          where: { shareId },
        });

        if (!files || files.length === 0) {
          throw new NotFoundException(`No files found for share ${shareId}`);
        }

        const archive = archiver("zip", {
          zlib: { level: parseInt(compressionLevel) },
        });

        archive.on("error", (err) => {
          this.logger.error("Archive error", err);
          reject(new InternalServerErrorException("Error creating ZIP file"));
        });

        resolve(archive);

        for (const fileRecord of files) {
          try {
            const remotePath = `${this.getFolderPath()}${shareId}/${fileRecord.name}`;
            const token = await this.getAccessToken();
            const downloadUrl = `${this.getDriveBaseUrl()}/root:/${remotePath}:/content`;

            const response = await fetch(downloadUrl, {
              method: "GET",
              headers: {
                Authorization: `Bearer ${token}`,
              },
            });

            if (response.ok && response.body) {
              const nodeStream = Readable.fromWeb(response.body as any);
              archive.append(nodeStream, { name: fileRecord.name });
            }
          } catch (error) {
            this.logger.error(
              `Error fetching file ${fileRecord.name} from SharePoint`,
              error,
            );
          }
        }

        archive.finalize();
      } catch (error) {
        this.logger.error("Error creating ZIP from SharePoint files", error);
        reject(new InternalServerErrorException("Error creating ZIP file"));
      }
    });
  }
}
