import { HttpClientResponse, SPHttpClient } from "@microsoft/sp-http";

class _AdventsService {
  public getGroupMember = async (context: any, groupName: string, userEmail: string): Promise<any> => {
    const resp = await context.spHttpClient.get(`${context.pageContext.web.absoluteUrl}/_api/web/SiteGroups/GetByName('${groupName}')/Users?$filter=Email eq '${userEmail}'`, SPHttpClient.configurations.v1);
    const response = this.checkStatus(resp);
    const data = await response.json();
    return data;
  }

  private checkStatus(response: HttpClientResponse): HttpClientResponse {
    if (response.status >= 200 && response.status < 300) {
      return response;
    } else {
      const error: any = new Error(response.statusText);
      error.response = response;
      throw error;
    }
  }
}

const AdventsService = new _AdventsService();
export default AdventsService;