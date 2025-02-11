import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { PDFDocument } from "pdf-lib";
// import { PageContext } from "@microsoft/sp-page-context";
import { AadHttpClient } from "@microsoft/sp-http";
import { ListViewCommandSetContext } from "@microsoft/sp-listview-extensibility";
import { IDocument } from "../models/IDocument";
import axios from "axios";

export class PDFService {
  // private _context: ListViewCommandSetContext;
  private static readonly _tenantId: string =
    "4c2c8480-d3f0-485b-b750-807ff693802f";
    private static readonly _tenantId1: string =
    "10ec8bae-a29d-4cfa-9565-0dc447757371";
  private static readonly _clientId: string =
    "fc8a45d7-4f78-4e55-a611-76b1424242db";
  private static readonly _clientSecret: string =
    "/m37itwI+twVvMoIMYjjA7lrASCih/F5ZU6nEdK3qG8=";
  private static readonly _clientUrl: string =
    "https://func-adp-thinktank-ecm-stg-002.azurewebsites.net";
  private static readonly _functionAppUrl: string =
    "https://func-adp-thinktank-ecm-stg-002.azurewebsites.net/api/ConvertAndMergePDF";
  private static readonly _tokenUrl: string = `https://login.microsoftonline.com/common/oauth2/v2.0/token`;
  public static readonly serviceKey: ServiceKey<PDFService> =
    ServiceKey.create<PDFService>("PDFService", PDFService);
  // constructor(serviceScope: ServiceScope) {
  //   this._context = serviceScope.consume(ListViewCommandSetContext.serviceKey);
  // }

  public async mergePDFs(
    pdfBuffers: ArrayBuffer[]
  ): Promise<Uint8Array> {
    const mergedPdf = await PDFDocument.create();
    for (const buffer of pdfBuffers) {
      const pdfDoc = await PDFDocument.load(buffer);
      const copiedPages = await mergedPdf.copyPages(
        pdfDoc,
        pdfDoc.getPageIndices()
      );
      copiedPages.forEach((page) => mergedPdf.addPage(page));
    }
    return mergedPdf.save();
  }

  // public async mergePDFs(
  //   documents: IDocument[],
  //   pdfBuffers: ArrayBuffer[],
  //   context: ListViewCommandSetContext
  // ): Promise<void> {
  //   const token: string = await this.getAccessToken();
  //   console.log(token);

  //   const files: Blob[] = pdfBuffers.map((pdfBuffer) => new Blob([pdfBuffer]));

  //   const formData = new FormData();
  //   for (let i = 0; i < files.length; i++) {
  //     formData.append(documents[i].name, files[i]);
  //   }

  //   const resonse = await axios.post(PDFService._functionAppUrl, formData, {
  //     headers: {
  //       Authorization: `Bearer ${token}`,
  //       "Content-Type": "multipart/form-data",
  //     },
  //   });
  //   // return context.aadHttpClientFactory
  //   //   .getClient(PDFService._clientUrl)
  //   //   .then((client: AadHttpClient) => {
  //   //     return client.get(
  //   //       PDFService._functionAppUrl,
  //   //       AadHttpClient.configurations.v1,
  //   //       { body: formData }
  //   //     );
  //   //   })
  //   //   .then((response) => response.json())
  //   //   .then((data) => console.log("API Data:", data))
  //   //   .catch((error) => console.error("Error:", error));
  // }

  public async getAccessToken(): Promise<string> {
    const requestBody = new URLSearchParams();
    requestBody.append("grant_type", "client_credentials");
    requestBody.append("client_id", PDFService._clientId);
    requestBody.append("client_secret", PDFService._clientSecret);
    requestBody.append("scope", `https://graph.microsoft.com/.default`);

    const response = await axios.post(PDFService._tokenUrl, requestBody, {
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
        "Access-Control-Allow-Origin": "*"
      },
    });
    console.log(response.data.access_token);
    return response.data.access_token;
  }
}
