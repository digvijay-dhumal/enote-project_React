import React, { useState, useEffect } from "react";
import Navbar from "../components/navbar";
import Footer from "../components/footer";
import { Label } from "@progress/kendo-react-labels";
import { DatePicker } from "@progress/kendo-react-dateinputs";
import { DropDownList } from "@progress/kendo-react-dropdowns";
import { useMsal, useAccount } from "@azure/msal-react";
import { Sidebar } from "../components/sidebar";
import { MenuNavContainer } from "../components/menu";
import DateObject from "react-date-object";
import "../styles/datagridpage.css";
import "../styles/forms.css";
import { API_BASE_URL, API_ENDPOINTS, API_COMMON_HEADERS, loginRequest } from "../config";
import { PageLoader } from "../components/pageLoader";
import { getAccessToken } from "../App";
import AdminDashboardGrid from "../components/DashBoardReporttableForAdmin";
import {
    Chart,
    ChartSeries,
    ChartSeriesItem,
    ChartTitle,
    ChartTooltip
} from "@progress/kendo-react-charts";
import "hammerjs";
import "bootstrap/dist/css/bootstrap.min.css";

const tooltipRender = props => {
    if (props.point) {
        const {
            category,
        } = props.point;
        return <span>{String(category)}</span>;
    }
};

const defaultDate = new Date();
defaultDate.setMonth(3); // April is month 3 (zero-based index)
defaultDate.setDate(1);

if (defaultDate.getMonth() >= 3) {
    defaultDate.setFullYear(defaultDate.getFullYear() - 1);
}

const AdminDashBoardChartReports = () => {
    const [dataResult, setDataResult] = useState([]);
    const [atrdataResult, setATRDataResult] = useState([]);
    const [isLoading, setIsLoading] = useState(true);
    const { accounts, instance } = useMsal();
    const account = useAccount(accounts[0] || {});
    const [fromDate, setFormDate] = useState(defaultDate);
    const [toDate, setToDate] = useState(new Date());
    const [selectedDepartment, setSelectedDepartment] = useState("");
    const [getDepartment, setGetDepartment] = useState([]);
    const [isSuperAdmin, setIsSuperAdmin] = useState(false);

    useEffect(() => {
        GetDepartments();
        getUserDepartment();
    }, []);

    useEffect(() => {
        getAdminType();
        fetchApiData();
        fetchApiATRRepostsData();
    }, [fromDate, toDate, selectedDepartment]);

    const GetDepartments = async () => {
        try {
            const accessToken = await getAccessToken({ ...loginRequest, account }, instance);

            const obj = await fetch(
                `${API_BASE_URL}${API_ENDPOINTS.GET_Departments}`, {
                method: "GET",
                headers: { ...API_COMMON_HEADERS, Authorization: `Bearer ${accessToken}` }
            }
            );
            const dep = await obj.json();

            setGetDepartment(dep);
        } catch (err) {
            console.log(err)
        }
    }
    const getUserDepartment = async () => {
        try {
            const accessToken = await getAccessToken({ ...loginRequest, account }, instance);

            const obj = await fetch(
                `${API_BASE_URL}${API_ENDPOINTS.GET_UserDetailsByPrincipalName(accounts[0].username)}`, {
                method: "GET",
                headers: { ...API_COMMON_HEADERS, Authorization: `Bearer ${accessToken}` }
            }
            );
            const departmentDetails = await obj.json();
            setSelectedDepartment(departmentDetails[0].department);
        } catch (error) {
            console.error("Error fetching user details:", error);
        }
        setIsLoading(false)
    };

    const getAdminType = async () => {
        try {
            const accessToken = await getAccessToken({ ...loginRequest, account }, instance);

            const obj = await fetch(
                `${API_BASE_URL}${API_ENDPOINTS.GET_AdminType}?userPrincipalName=${accounts[0].username}`, {
                method: "GET",
                headers: { ...API_COMMON_HEADERS, Authorization: `Bearer ${accessToken}` }
            }
            )
            const adminType = await obj.text();
            setIsSuperAdmin(adminType === "Super Admin");
        } catch (err) {
            console.log(err)
        }
    }

    const fetchApiData = async () => {
        const colorCodes = [
            "#7731a3", // Deep Purple
            "rgb(254, 215, 76)", // Yellow
            "#ba3294", // Purple
            "rgb(78, 185, 167)", // Turquoise
            "#d92b2b", // Red
            "rgb(12, 77, 162)", // Blue
            "#4dc313", // Green
            "#1e6bb5", // Dark Blue
            "rgb(255, 165, 0)" // Orange
        ];

        if (fromDate && toDate && (selectedDepartment)) {
            try {
                const accessToken = await getAccessToken({ ...loginRequest, account }, instance);
                const Params = {
                    "Department": selectedDepartment.replace(/&/g, "%26"),
                    "fromDate": fromDate ? new DateObject(new Date(fromDate)).format("DD-MM-YYYY") : null,
                    "ToDate": toDate ? new DateObject(new Date(toDate)).format("DD-MM-YYYY") : null,
                }
                const obj = await fetch(
                    `${API_BASE_URL}${API_ENDPOINTS.GET_ReportChartsForAdmin}?Department=${Params.Department}&fromDate=${Params.fromDate}&ToDate=${Params.ToDate}`, {
                    method: "GET",
                    headers: { ...API_COMMON_HEADERS, Authorization: `Bearer ${accessToken}` },
                }
                );
                const res = await obj.json();
                if (res) {
                    res.lstNoteChart = res.lstNoteChart?.map((obj, index) => { return { ...obj, color: colorCodes[index], status: `${obj.status}: ${obj.count}` } })
                    setDataResult(res.lstNoteChart)
                }
            }
            catch (err) {
                console.log(err)
            }
        }
    }

    // ATR Reports
    const fetchApiATRRepostsData = async () => {
        const colorCodes = [
            "#7731a3", // Deep Purple
            "rgb(254, 215, 76)", // Yellow
            "#ba3294", // Purple
            "rgb(78, 185, 167)", // Turquoise
            "#d92b2b", // Red
            "rgb(12, 77, 162)", // Blue
            "#4dc313", // Green
            "#1e6bb5", // Dark Blue
            "rgb(255, 165, 0)" // Orange
        ];

        if (fromDate && toDate && selectedDepartment) {
            try {
                const accessToken = await getAccessToken({ ...loginRequest, account }, instance);
                const Params = {
                    "Department": selectedDepartment.replace(/&/g, "%26"),
                    "fromDate": fromDate ? new DateObject(new Date(fromDate)).format("DD-MM-YYYY") : null,
                    "ToDate": toDate ? new DateObject(new Date(toDate)).format("DD-MM-YYYY") : null,
                }
                const obj = await fetch(
                    `${API_BASE_URL}${API_ENDPOINTS.Get_ATRReport}?Department=${Params.Department}&fromDate=${Params.fromDate}&ToDate=${Params.ToDate}`, {
                    method: "GET",
                    headers: { ...API_COMMON_HEADERS, Authorization: `Bearer ${accessToken}` },
                }
                );
                const res = await obj.json();
                if (res) {
                    console.log(res, "res")
                    res.lstNoteChart = res.lstNoteChart?.map((obj, index) => { return { ...obj, color: colorCodes[index], status: `${obj.status}: ${obj.count}` } })
                    setATRDataResult(res.lstNoteChart)
                }
            }
            catch (err) {
                console.log(err)
            }
        }
    }
    const handelSelectedFromDate = (e) => {
        setFormDate(e.target.value);
    }

    const handelSelectedToDate = (e) => {
        setToDate(e.target.value);
    }

    return (
        <div>
            <Navbar header="IB Smart Office - eNote" />
            <Sidebar />
            {/* Added top navigation for menu --> 23-09 */}
            <MenuNavContainer isMenuPage />
            <div className="container largeScreen cstGridContainer datagridpage ">
                <div className="SectionHeads row mobileSectionHeads">Dashboard Reports</div>
                <div className="row">
                    <div className="col-md-4">
                        <Label className="k-form-label" style={{ fontSize: "14px" }}>From Date:</Label>
                        <DatePicker
                            max={new Date()}
                            placeholder="From Date..."
                            onChange={handelSelectedFromDate}
                            format={"dd-MM-yyyy"}
                            value={fromDate}
                        />

                    </div>
                    <div className="col-md-4">
                        <Label className="k-form-label" style={{ fontSize: "14px" }}>To Date:</Label>
                        <DatePicker
                            min={new Date(fromDate)}
                            max={new Date()}
                            placeholder="To Date..."
                            onChange={handelSelectedToDate}
                            format={"dd-MM-yyyy"}
                            value={toDate}
                        />
                    </div>
                    {isSuperAdmin && (
                        <div className="col-md-4">
                            <Label className="k-form-label" style={{ fontSize: "14px" }}>Select Department:</Label>
                            <DropDownList
                                data={getDepartment?.map(x => x.departmentName)}
                                value={selectedDepartment}
                                onChange={(e) =>
                                    setSelectedDepartment(e.target.value)
                                }
                            />
                        </div>
                    )}
                </div>
                {/* Add a section for displaying headings and allow the user to switch between them*/}
                {isLoading ? (
                    <div>
                        {/* <PageLoader /> */}
                    </div>
                ) : (
                    <>
                        <div className="SectionRowForPowerBIReports row">
                            <div className="col-md-6 custom-pie-chart-seperateline ">
                                <div className="row">
                                    <div className="enote-chartContainer">
                                        <Chart>
                                            <ChartTitle text="Notes - Status" color="#7731a3" />
                                            <ChartSeries>
                                                <ChartSeriesItem
                                                    type="pie"
                                                    data={dataResult}
                                                    field="count"
                                                    categoryField="status"
                                                    tooltip={{
                                                        visible: true,
                                                    }}
                                                />

                                            </ChartSeries>
                                            <ChartTooltip opacity={1} visible={true} render={tooltipRender} />
                                        </Chart>
                                    </div>
                                </div>
                            </div>
                            <div className="col-md-6">
                                <div className="row">
                                    <div className="enote-chartContainer">
                                        <Chart>
                                            <ChartTitle text="ATR - Status" color="#7731a3" />
                                            <ChartSeries>
                                                <ChartSeriesItem
                                                    type="pie"
                                                    data={atrdataResult}
                                                    field="count"
                                                    categoryField="status"
                                                    tooltip={{
                                                        visible: true,
                                                    }}
                                                />

                                            </ChartSeries>
                                            <ChartTooltip opacity={1} visible={true} render={tooltipRender} />
                                        </Chart>
                                    </div>
                                </div>
                            </div>
                        </div>
                        <div className="SectionHeadsDashBoard row">Notes - Pending With</div>
                        <div className="row">
                            <div className="col">
                                <AdminDashboardGrid departmentName={selectedDepartment} fromDate={fromDate} toDate={toDate} />
                            </div>
                        </div>
                    </>
                )}
            </div>
            <div className="pgFooterContainer">
                <Footer />
            </div>
        </div>
    );
};

export default AdminDashBoardChartReports;