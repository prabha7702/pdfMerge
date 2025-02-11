import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { PageContext } from "@microsoft/sp-page-context";

export class SharepointService {
  public static readonly serviceKey: ServiceKey<SharepointService> =
    ServiceKey.create<SharepointService>("SPServiceHttp", SharepointService);
  private _spHttpClient: SPHttpClient;
  private _context: PageContext;
  private static _maxRetries = 3;
  private static _retryDelay = 1000;
  constructor(serviceScope: ServiceScope) {
    serviceScope.whenFinished(() => {
      this._spHttpClient = serviceScope.consume(SPHttpClient.serviceKey);
      this._context = serviceScope.consume(PageContext.serviceKey);
    });
  }

  public async getFileRef(id: number): Promise<string> {
    const response = await this._spHttpClient
      .get(
        this._context.web.absoluteUrl.concat(
          `/_api/web/lists/getbytitle('Documents')/items(${id})?$select=FileRef`
        ),
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => response.json());
    return response.FileRef;
  }

  public async getFileContent(fileref: string): Promise<ArrayBuffer> {
    const response: SPHttpClientResponse = await this._spHttpClient.get(
      this._context.web.absoluteUrl.concat(
        `/_api/web/getfilebyserverrelativeurl('${fileref}')/$value`
      ),
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: "application/octet-stream",
        },
      }
    );
    return response.arrayBuffer();
  }

  public async checkFileExist(folderPath:string,filename: string): Promise<boolean> {
    const response: SPHttpClientResponse = await this._spHttpClient.get(
      this._context.web.absoluteUrl.concat(
        `/_api/web/getfolderbyserverrelativepath(decodedUrl='${folderPath}')/files('${filename}')?$select=Exists`
      ),
      SPHttpClient.configurations.v1
    );
    return response.ok;
  }

  //   public async getFileContent(id: number): Promise<any> {
  //     return this.getFileRef(id)
  //       .then((response: any) => response.FileRef)
  //       .then((fileref: string) => {
  //         return this._spHttpClient
  //           .get(
  //             this._context.web.absoluteUrl.concat(
  //               `/_api/web/lists/getbytitle('Documents')/items(${id})/file/$value`
  //             ),
  //             SPHttpClient.configurations.v1,
  //             {
  //               headers: {
  //                 Accept: "application/octet-stream"
  //               },
  //             }
  //           )
  //           .then((response: SPHttpClientResponse) => {
  //             console.log(`Response: ${response}`);
  //             return response.arrayBuffer()
  //       });
  //       });
  //   }

  public async uploadFile(
    fileContent: Uint8Array,
    filename: string,
    folderPath: string
  ): Promise<JSON> {
    const response: SPHttpClientResponse = await this._spHttpClient.post(
      this._context.web.absoluteUrl.concat(
        `/_api/web/getfolderbyserverrelativeurl('${folderPath}')/files/add(url='${filename}.pdf')`
      ),
      SPHttpClient.configurations.v1,
      {
        headers: {
          "Accept": "application/json;odata=verbose",
          "Content-Type": "application/pdf",
        },
        body: fileContent,
      }
    );
    return response.json();
  }

  public async deleteFile(fileref: string, retries: number = 0): Promise<JSON> {
    const response: SPHttpClientResponse = await this._spHttpClient.post(
      this._context.web.absoluteUrl.concat(
        `/_api/web/getfolderbyserverrelativeurl('${fileref}')`
      ),
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: "application/json;odata=verbose",
          "IF-MATCH": "*",
          "X-HTTP-Method": "DELETE",
        },
      }
    );
    if (response.ok) {
      return response.json();
    } else {
      if (retries < SharepointService._maxRetries) {
        await new Promise((resolve) =>
          setTimeout(resolve, SharepointService._retryDelay)
        );
        return this.deleteFile(fileref, retries + 1);
      } else {
        throw new Error(response.statusText);
      }
    }
  }
}
