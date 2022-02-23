import React, { Component } from 'react';
import { Dialog, Flex, CloseIcon, FormInput, FormDropdown, SyncIcon, Button } from '@fluentui/react-northstar';
import "../scss/UpdateIncident.module.scss";
import { Col, Row } from 'react-bootstrap';
import { Client } from "@microsoft/microsoft-graph-client";
import * as graphConfig from '../common/graphConfig';
import siteConfig from '../config/siteConfig.json';
import * as constants from '../common/Constants';
import CommonService from "../common/CommonService";

export interface IListItem {
    itemId: string;
    incidentId: string;
    incidentName: string;
    incidentCommander: string;
    status: string;
    location: string;
    startDate: string;
    startDateUTC: string;
    createdDate: string;
}

export interface IUpdateIncidentProps {
    openPopup: boolean;
    closePopup: (isRefreshNeeded: boolean) => void;
    incidentData: IListItem;
    graph: Client;
    tenantName: string;
    siteId: string;
    showMessageBar(message: string, type: string): void;
    localeStrings: any;
};

export interface IUpdateIncidentState {
    isDesktop: boolean;
    statusOptions: [];
    selectedStatus: string;
};

export default class UpdateIncident extends Component<IUpdateIncidentProps, IUpdateIncidentState> {
    constructor(props: IUpdateIncidentProps) {
        super(props);
        this.state = {
            isDesktop: window.innerWidth >= constants.mobileWidth ? true : false,
            statusOptions: [],
            selectedStatus: ""
        };
    }

    // initialize data service
    private dataService = new CommonService();

    public async componentDidMount() {
        await this.getStatusOptions();
    }

    // get dropdown options for status
    public async getStatusOptions() {
        try {
            const incStatusGraphEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}${graphConfig.listsGraphEndpoint}/${siteConfig.incStatusList}/items?$expand=fields&$Top=5000`;
            const statusOptions = await this.dataService.getDropdownOptions(incStatusGraphEndpoint, this.props.graph);
            this.setState({
                statusOptions: statusOptions
            })
        } catch (error) {
            console.error(
                constants.errorLogPrefix + "_UpdateIncident_GetStatusOptions \n",
                JSON.stringify(error)
            );
        }
    }

    // on status dropdown value change
    private onStatusChange = (event: any, selectedValue: any) => {
        this.setState({
            selectedStatus: selectedValue.value
        })
    }

    // on update button click
    private updateStatus = async () => {
        try {
            const updatedStatus = {
                IncidentStatus: this.state.selectedStatus !== "" ? this.state.selectedStatus : this.props.incidentData.status
            }
            let graphEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}/lists/${siteConfig.incidentsList}/items/${this.props.incidentData.itemId}/fields`;

            // let service = new CommonService();
            const updatedItem = await this.dataService.updateItemInList(graphEndpoint, this.props.graph, updatedStatus);
            if (updatedItem) {
                console.log(constants.infoLogPrefix + "Incident Updated");
                this.props.closePopup(true);
                this.props.showMessageBar(this.props.localeStrings.updateStatusSuccessMessage, constants.messageBarType.success);
            }
        } catch (error) {
            console.error(
                constants.errorLogPrefix + "_UpdateIncident_UpdateStatus \n",
                JSON.stringify(error)
            );
        }
    }

    render() {
        return (
            <div>
                <Dialog
                    open={this.props.openPopup}
                    closeOnOutsideClick={false}
                    id="incident-popup"
                    content={<>
                        <Flex space="between" id="incident-popup-header">
                            <div className="popup-header-text">
                                {this.props.localeStrings.manageIncFormTitle}
                            </div>
                            <CloseIcon onClick={() => { this.props.closePopup(false); }} id="popup-header-close" />
                        </Flex>
                        <div className="incident-popup-body">
                            <Row xs={1} sm={2} md={3}>
                                <Col md={4} sm={6} xs={12}>
                                    <div className="popup-grid-item">
                                        <FormInput
                                            label={this.props.localeStrings.fieldIncidentId}
                                            type="text"
                                            placeholder={this.props.localeStrings.phIncidentId}
                                            fluid={true}
                                            value={this.props.incidentData ? this.props.incidentData.incidentId : ""}
                                            disabled
                                            id="popup-text-field"
                                        />
                                    </div>
                                    <div className="popup-grid-item">
                                        <FormDropdown
                                            label={{ content: this.props.localeStrings.fieldIncidentStatus, required: true, className: "status-dd-label" }}
                                            placeholder={this.props.localeStrings.phIncidentStatus}
                                            items={this.state.statusOptions}
                                            fluid={true}
                                            value={this.state.selectedStatus !== "" ? this.state.selectedStatus : this.props.incidentData.status}
                                            onChange={this.onStatusChange}
                                            id="popup-dropdown"
                                        />
                                    </div>
                                </Col>
                                <Col md={4} sm={6} xs={12}>
                                    <div className="popup-grid-item">
                                        <FormInput
                                            label={{ content: this.props.localeStrings.fieldIncidentName }}
                                            placeholder={this.props.localeStrings.phIncidentName}
                                            fluid={true}
                                            value={this.props.incidentData ? this.props.incidentData.incidentName : ""}
                                            disabled
                                            id="popup-text-field"
                                        />
                                    </div>
                                    <div className="popup-grid-item">
                                        <FormInput
                                            label={this.props.localeStrings.fieldLocation}
                                            type="text"
                                            placeholder={this.props.localeStrings.phLocation}
                                            fluid={true}
                                            value={this.props.incidentData ? this.props.incidentData.location : ""}
                                            disabled
                                            id="popup-text-field"
                                        />
                                    </div>
                                </Col>
                                <Col md={4} sm={6} xs={12}>
                                    <div className="popup-grid-item">
                                        <FormInput
                                            label={this.props.localeStrings.fieldIncidentCommander}
                                            type="text"
                                            placeholder={this.props.localeStrings.phIncidentCommander}
                                            fluid={true}
                                            value={this.props.incidentData ? this.props.incidentData.incidentCommander : ""}
                                            disabled
                                            id="popup-text-field"
                                        />
                                    </div>
                                    <div className="popup-grid-item">
                                        <FormInput
                                            label={this.props.localeStrings.fieldStartDate}
                                            type="text"
                                            placeholder={this.props.localeStrings.phStartDate}
                                            fluid={true}
                                            defaultValue={this.props.incidentData.startDate}
                                            disabled
                                            id="popup-date-field"
                                        />
                                    </div>
                                </Col>
                            </Row>
                        </div>
                        <Flex hAlign={this.state.isDesktop ? "end" : "center"} gap="gap.small" id="popup-btn-area">
                            <Button
                                icon={<CloseIcon />}
                                content={this.props.localeStrings.btnClose}
                                iconPosition="before"
                                id="popup-close-btn"
                                title={this.props.localeStrings.btnClose}
                                onClick={() => { this.props.closePopup(false); }}
                            />
                            <Button
                                icon={<SyncIcon />}
                                content={this.props.localeStrings.btnUpdateInc}
                                iconPosition="before"
                                primary
                                onClick={this.updateStatus}
                                id="popup-update-btn"
                                title={this.props.localeStrings.btnUpdateInc}
                                fluid={this.state.isDesktop ? false : true}
                            />
                        </Flex>
                    </>}
                />
            </div>
        )
    }
}
