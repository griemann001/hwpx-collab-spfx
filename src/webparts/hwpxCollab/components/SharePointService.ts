/**
 * ============================================================
 * SharePointService.ts — SharePoint 문서 라이브러리 연동
 * ============================================================
 */

const FOLDER_PATH = 'Shared Documents/HwpxCollab';

export interface IHwpxFileInfo {
  id: number;
  name: string;
  serverRelativeUrl: string;
  modifiedDate: string;
  modifiedBy: string;
  createdByEmail: string;
  size: number;
}

export class SharePointService {
  private _spHttpClient: any;
  private _siteUrl: string;

  constructor(spHttpClient: any, siteUrl: string) {
    this._spHttpClient = spHttpClient;
    this._siteUrl = siteUrl;
  }

  public async getFileList(): Promise<IHwpxFileInfo[]> {
    try {
      // 1차: Files 목록 + AuthorId (ListItemAllFields에서 함께 조회)
      const filesUrl = `${this._siteUrl}/_api/web/GetFolderByServerRelativeUrl('${FOLDER_PATH}')/Files?$select=Name,ServerRelativeUrl,TimeLastModified,Length,ListItemAllFields/Id,ListItemAllFields/AuthorId,ModifiedBy/Title&$expand=ListItemAllFields,ModifiedBy&$orderby=TimeLastModified desc`;
      console.log('[SP] 파일 목록 조회:', filesUrl);
      const filesResp = await this._spHttpClient.get(filesUrl, this._spHttpClient.constructor.configurations.v1);
      if (!filesResp.ok) {
        console.error('[SP] 파일 목록 조회 실패:', filesResp.status, await filesResp.text());
        return [];
      }
      const filesData = await filesResp.json();
      console.log('[SP] 파일 목록 응답:', filesData);

      const hwpxFiles = (filesData.value || []).filter((f: any) => f.Name && f.Name.toLowerCase().endsWith('.hwpx'));
      if (hwpxFiles.length === 0) return [];

      // 2차: AuthorId → siteusers/getbyid() 로 이메일 조회 (중복 AuthorId 캐싱)
      const uniqueAuthorIds: number[] = Array.from(new Set(
        hwpxFiles
          .map((f: any) => f.ListItemAllFields && f.ListItemAllFields.AuthorId)
          .filter((id: any) => typeof id === 'number')
      ));
      const authorMap: Record<number, string> = {};
      await Promise.all(uniqueAuthorIds.map(async (authorId: number) => {
        try {
          const userUrl = `${this._siteUrl}/_api/web/siteusers/getbyid(${authorId})?$select=Email,LoginName`;
          const userResp = await this._spHttpClient.get(userUrl, this._spHttpClient.constructor.configurations.v1);
          if (userResp.ok) {
            const userData = await userResp.json();
            authorMap[authorId] = userData.Email || '';
            console.log(`[SP] Author(${authorId}) email:`, userData.Email, 'login:', userData.LoginName);
          } else {
            console.warn(`[SP] siteusers getbyid(${authorId}) 실패:`, userResp.status);
          }
        } catch (e) {
          console.warn(`[SP] Author(${authorId}) 조회 실패:`, e);
        }
      }));
      console.log('[SP] Author 이메일 맵:', authorMap);

      const files: IHwpxFileInfo[] = hwpxFiles.map((f: any) => {
        const itemId: number = f.ListItemAllFields ? f.ListItemAllFields.Id : 0;
        const authorId: number = f.ListItemAllFields ? f.ListItemAllFields.AuthorId : 0;
        return {
          id: itemId,
          name: f.Name,
          serverRelativeUrl: f.ServerRelativeUrl,
          modifiedDate: f.TimeLastModified ? new Date(f.TimeLastModified).toISOString().split('T')[0] : '',
          modifiedBy: f.ModifiedBy ? f.ModifiedBy.Title : '',
          createdByEmail: authorMap[authorId] || '',
          size: f.Length || 0,
        };
      });

      console.log('[SP] hwpx 파일 수:', files.length);
      if (files.length > 0) {
        console.log('[SP] 첫 파일 createdByEmail:', files[0].createdByEmail);
      }
      return files;
    } catch (err) {
      console.error('[SP] 파일 목록 조회 에러:', err);
      return [];
    }
  }

  public async uploadFile(fileName: string, fileBuffer: ArrayBuffer): Promise<IHwpxFileInfo | null> {
    const url = `${this._siteUrl}/_api/web/GetFolderByServerRelativeUrl('${FOLDER_PATH}')/Files/add(url='${encodeURIComponent(fileName)}',overwrite=true)`;

    try {
      console.log('[SP] 파일 업로드:', url);
      const response = await this._spHttpClient.post(
        url,
        this._spHttpClient.constructor.configurations.v1,
        {
          body: fileBuffer,
          headers: {
            'Content-Type': 'application/octet-stream',
          },
        }
      );

      if (!response.ok) {
        const errText = await response.text();
        console.error('[SP] 파일 업로드 실패:', response.status, errText);
        return null;
      }

      const data = await response.json();
      console.log('[SP] 업로드 응답:', data);
      return {
        id: 0,
        name: data.Name || fileName,
        serverRelativeUrl: data.ServerRelativeUrl || '',
        modifiedDate: new Date().toISOString().split('T')[0],
        modifiedBy: '',
        createdByEmail: '',
        size: fileBuffer.byteLength,
      };
    } catch (err) {
      console.error('[SP] 파일 업로드 에러:', err);
      return null;
    }
  }

  public async downloadFile(serverRelativeUrl: string): Promise<ArrayBuffer | null> {
    const url = `${this._siteUrl}/_api/web/GetFileByServerRelativeUrl('${serverRelativeUrl}')/$value`;

    try {
      console.log('[SP] 파일 다운로드:', url);
      const response = await this._spHttpClient.get(
        url,
        this._spHttpClient.constructor.configurations.v1
      );

      if (!response.ok) {
        console.error('[SP] 파일 다운로드 실패:', response.status);
        return null;
      }

      const buffer = await response.arrayBuffer();
      console.log('[SP] 다운로드 완료:', buffer.byteLength, 'bytes');
      return buffer;
    } catch (err) {
      console.error('[SP] 파일 다운로드 에러:', err);
      return null;
    }
  }

  public async saveFile(fileName: string, fileBuffer: ArrayBuffer): Promise<boolean> {
    const result = await this.uploadFile(fileName, fileBuffer);
    return result !== null;
  }

  public async deleteFile(serverRelativeUrl: string): Promise<boolean> {
    const url = `${this._siteUrl}/_api/web/GetFileByServerRelativeUrl('${serverRelativeUrl}')`;

    try {
      const response = await this._spHttpClient.post(
        url,
        this._spHttpClient.constructor.configurations.v1,
        {
          headers: {
            'X-HTTP-Method': 'DELETE',
            'IF-MATCH': '*',
          },
        }
      );
      return response.ok;
    } catch (err) {
      console.error('[SP] 파일 삭제 에러:', err);
      return false;
    }
  }
}