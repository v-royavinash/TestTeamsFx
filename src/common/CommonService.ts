import { Client } from "@microsoft/microsoft-graph-client";
import * as constants from './Constants';
import moment from "moment";

export interface IListItem {
    itemId: string;
    incidentId: string;
    incidentName: string;
    incidentCommander: string;
    status: string;
    location: string;
    startDate: string;
    startDateUTC: string;
    modifiedDate: string;
    teamWebURL: string;
}
export default class CommonService {

    //#region Dashboard Methods

    // get data to show on the Dashboard
    public async getDashboardData(graphEndpoint: any, graph: Client): Promise<any> {
        try {

            const incidentsData = await graph.api(graphEndpoint).get();

            // Prepare the output array
            var formattedIncidentsData: Array<IListItem> = new Array<IListItem>();

            // Map the JSON response to the output array
            incidentsData.value.forEach((item: any) => {
                formattedIncidentsData.push({
                    itemId: item.fields.id,
                    incidentId: item.fields.IncidentId,
                    incidentName: item.fields.IncidentName,
                    incidentCommander: item.fields.IncidentCommander,
                    status: item.fields.IncidentStatus,
                    location: item.fields.Location,
                    startDate: moment(item.fields.StartDateTime).format("DD MMM, yyyy"),
                    startDateUTC: new Date(item.fields.StartDateTime).toISOString().slice(0, new Date(item.fields.StartDateTime).toISOString().length - 1),
                    modifiedDate: item.fields.Modified,
                    teamWebURL: item.fields.TeamWebURL
                });
            });
            return formattedIncidentsData;

        } catch (error) {
            console.error(
                constants.errorLogPrefix + "_CommonService_GetDashboardData \n",
                JSON.stringify(error)
            );
        }
    }
    //#endregion

    //#region Create Incident Methods

    // get dropdown options for Incident Type, Status and Role Assignments
    public async getDropdownOptions(graphEndpoint: any, graph: Client): Promise<any> {
        try {
            const listData = await graph.api(graphEndpoint).get();

            // Prepare the output array
            const drpdwnOptions: Array<any> = new Array<any>();

            // Map the JSON response to the output array
            listData.value.forEach((item: any) => {
                drpdwnOptions.push(
                    item.fields.Title
                );
            });
            return drpdwnOptions;
        } catch (error) {
            console.error(
                constants.errorLogPrefix + "_CommonService_GetDropdownOptions \n",
                JSON.stringify(error)
            );
        }
    }

    // Generic method to update item to SharePoint list
    public async updateItemInList(graphEndpoint: any, graph: Client, listItemObj: any): Promise<any> {
        return await graph.api(graphEndpoint).update(listItemObj);
    }

    // get incident details based on incident name
    public async getExistingIncident(graphEndpoint: any, graph: Client): Promise<any> {

        return await graph.api(graphEndpoint)
            .header('Prefer', 'HonorNonIndexedQueriesWarningMayFailRandomly')
            .get();
    }

    // create channel
    public async createChannel(graphEndpoint: any, graph: Client, channelObj: any): Promise<any> {
        return await graph.api(graphEndpoint).post(JSON.stringify(channelObj));
    }

    // generic method for a POST graph query
    public async sendGraphPostRequest(graphEndpoint: any, graph: Client, requestObj: any): Promise<any> {
        return await graph.api(graphEndpoint).post(requestObj);
    }

    // generic method for a PUT graph query
    public async sendGraphPutRequest(graphEndpoint: any, graph: Client, requestObj: any): Promise<any> {
        return await graph.api(graphEndpoint).put(requestObj);
    }

    // generic method for a PATCH graph query
    public async sendGraphPatchRequest(graphEndpoint: any, graph: Client, requestObj: any): Promise<any> {
        return await graph.api(graphEndpoint).patch(requestObj);
    }

    // generic method for a Delete graph query
    public async sendGraphDeleteRequest(graphEndpoint: any, graph: Client): Promise<any> {
        return await graph.api(graphEndpoint).delete();
    }

    //#endregion

    //#region Common Methods

    // Get tenant name
    public async getTenantDetails(graphEndpoint: any, graph: Client): Promise<any> {
        try {
            const domains = await graph.api(graphEndpoint).get();

            let domainName = "";
            domains.value.forEach((element: any) => {
                element.verifiedDomains.forEach((vDomains: any) => {
                    if (vDomains.isInitial) {
                        if (vDomains.name.indexOf('.onmicrosoft.com') > -1) {
                            domainName = vDomains.name.split('.onmicrosoft.com')[0];
                        }
                    }
                });
            });

            return domainName;
        } catch (error) {
            console.error(
                constants.errorLogPrefix + "_CommonService_GetTenantDetails \n",
                JSON.stringify(error)
            );
        }
    }

    // this is generic method to return the graph data based on input graph endpoint
    public async getGraphData(graphEndpoint: any, graph: Client): Promise<any> {
        return await graph.api(graphEndpoint).get();
    }
    //#endregion
}