import React, { useState, useEffect } from "react";
// import Kendo components
import {
  Grid,
  GridColumn as Column,
  GridToolbar,
} from "@progress/kendo-react-grid";
import { Link } from "react-router-dom";
import { Button } from "@progress/kendo-react-buttons";
import { setGroupIds } from "@progress/kendo-react-data-tools";
import { getter } from "@progress/kendo-react-common";
import { orderBy } from "@progress/kendo-data-query";
import { process } from "@progress/kendo-data-query";
import { Input } from "@progress/kendo-react-inputs";
// import external components
import { useMsal, useAccount } from "@azure/msal-react";
import { CSVLink } from "react-csv";
import DateObject from "react-date-object";
import { ColumnMenu } from "./custom-cells";
import Navbar from "../components/navbar";
import Footer from "../components/footer";
import { useParams } from "react-router";
// import internal components
import { Sidebar } from "../components/sidebar";
import { MenuNavContainer } from "../components/menu";
import {
  API_BASE_URL,
  API_ENDPOINTS,
  loginRequest,
  API_COMMON_HEADERS,
  IB_eReceptionDomain_URL,
} from "../config";
import { getAccessToken } from "../App";
import { Dialog, DialogActionsBar } from "@progress/kendo-react-dialogs";
// import { API_BASE_URL, API_ENDPOINTS, API_COMMON_HEADERS, loginRequest, IB_eReceptionDomain_URL } from "../config";
// import styles
import "../styles/datagridpage.css";
import "../styles/forms.css";

const initialSort = [
  {
    field: "modifiedDate", //noteId
    dir: "desc",
  },
];
const DATA_ITEM_KEY = "noteId";
const SELECTED_FIELD = "Action"; 
const initialDataState = {
  take: 10,
  skip: 0,
  group: [],
};

const processWithGroups = (data, dataState) => {
  const newDataState = process(data, dataState);
  setGroupIds({
    data: newDataState.data,
    group: dataState.group,
  });
  return newDataState;
};

const Views = () => {
  const idGetter = getter(DATA_ITEM_KEY);
  const { id } = useParams();
  const [filterValue, setFilterValue] = React.useState("");
  const [filteredData, setFilteredData] = React.useState();
  const [currentSelectedState, setCurrentSelectedState] = React.useState({});
  const [dataState, setDataState] = React.useState(initialDataState);
  const [data, setData] = React.useState(filteredData);
  const [apiData, setApiData] = React.useState([]);
  const [dataResult, setDataResult] = useState([]);
  const [isLoading, setIsLoading] = useState(true);
  const [currentHeading, setCurrentHeading] = useState();
  const [selectedView, setSelectedView] = useState("");
  const { accounts, instance } = useMsal();
  const account = useAccount(accounts[0] || {});
  const [filterErrorMessage, setFilterErrorMessage] = useState("");
  const [enumObj, setEnumObj] = useState();

   // Added states for delete functionality
  const [showDeleteAlert, setShowDeleteAlert] = useState(false);
  const [dataItemToDelete, setDataItemToDelete] = useState(null);
  const [showNotification, setShowNotification] = useState(false);
  const [alertMsg, setAlertMsg] = useState("");

  const getRequestByStatus = async (noteId) => {
    const accessToken = await getAccessToken(
      { ...loginRequest, account },
      instance
    );
    //get enum objects
    const dropdowns = await fetch(
      `${API_BASE_URL}${API_ENDPOINTS.GET_DROPDOWNDATA}`,
      {
        method: "GET",
        headers: {
          ...API_COMMON_HEADERS,
          Authorization: `Bearer ${accessToken}`,
        },
      }
    );
    const enumsObj = await dropdowns.json();

    setEnumObj(enumsObj);

    switch (noteId) {
      case "All Requests":
        fetchApiData(0, API_ENDPOINTS.eNote_GetRequestsRoleBased, accessToken);
        setSelectedView(noteId);
        break;
      case "Draft Requests":
        fetchApiData(
          enumsObj.NoteStatus.find((x) => x.dValue === "Draft").id,
          API_ENDPOINTS.eNote_GetRequestsRoleBased,
          accessToken
        );
        setSelectedView(noteId);
        break;
      case "In Progress":
        fetchApiData(
          enumsObj.NoteStatus.find((x) => x.dValue === "Pending").id,
          API_ENDPOINTS.eNote_GetRequestsRoleBased,
          accessToken
        );
        setSelectedView(noteId);
        break;
      case "All Approved":
        fetchApiData(
          enumsObj.NoteStatus.find((x) => x.dValue === "Approved").id,
          API_ENDPOINTS.eNote_GetRequestsRoleBased,
          accessToken
        );
        setSelectedView(noteId);
        break;
      case "All Rejected":
        fetchApiData(
          enumsObj.NoteStatus.find((x) => x.dValue === "Rejected").id,
          API_ENDPOINTS.eNote_GetRequestsRoleBased,
          accessToken
        );
        setSelectedView(noteId);
        break;
      case "Noted Notes":
        fetchApiData(
          "noted",
          API_ENDPOINTS.eNote_GetNotedNoteRequests,
          accessToken
        );
        setSelectedView(noteId);
        break;
      default:
        fetchApiData(0, API_ENDPOINTS.eNote_GetRequestsRoleBased, accessToken);
        setSelectedView("All Requests");
        return 0;
    }
  };

  const switchNoteId = async (id) => {
    setCurrentHeading(id);
    setIsLoading(true);
    getRequestByStatus(id);
  };
  useEffect(() => {
    switchNoteId(id); // Set the default noteId
  }, [id]);

  // const fetchApiData = async (status, endPoint, accessToken) => {
  //   setFilterValue("");
  //   const params =
  //     status === "noted"
  //       ? { CreatedBy: accounts[0].username }
  //       : { Status: status, CreatedBy: accounts[0].username };
  //   try {
  //     const response = await fetch(`${API_BASE_URL}${endPoint}`, {
  //       method: "POST",
  //       body: JSON.stringify(params),
  //       headers: {
  //         ...API_COMMON_HEADERS,
  //         Authorization: `Bearer ${accessToken}`,
  //       },
  //     });

  //     if (response.ok) {
  //       const apiOutput = await response.json();
  //       const resData = apiOutput.pendingNoteList;
  //       const apiData = orderBy(resData, initialSort);

  //       if (Array.isArray(apiData)) {
  //         setApiData(apiData);
  //         setFilteredData(apiData);
  //         setData(apiData);
  //         setDataResult(process(apiData, dataState));
  //         setDataState({ ...dataState, total: apiData.length });
  //       } else {
  //         console.error("Error fetching data: Invalid API response", apiData);
  //         setFilteredData([]);
  //         setDataResult(process([], dataState));
  //       }
  //     } else {
  //       console.error("Error fetching data:", response.statusText);
  //     }
  //   } catch (error) {
  //     console.error("Error fetching data:", error);
  //   } finally {
  //     setIsLoading(false);
  //   }
  // };
const fetchApiData = async (status, endPoint, accessToken) => {
  setFilterValue("");
  const params =
    status === "noted"
      ? { CreatedBy: accounts[0].username }
      : { Status: status, CreatedBy: accounts[0].username };
  try {
    const response = await fetch(`${API_BASE_URL}${endPoint}`, {
      method: "POST",
      body: JSON.stringify(params),
      headers: {
        ...API_COMMON_HEADERS,
        Authorization: `Bearer ${accessToken}`,
      },
    });

    if (response.ok) {
      const apiOutput = await response.json();
      const resData = apiOutput.pendingNoteList;
      const apiData = orderBy(resData, initialSort);

      if (Array.isArray(apiData)) {
        // 🔹 normalize the fields before saving
        const normalizedData = apiData.map((item) => ({
          ...item,
          subject: item.subject?.trim() ? item.subject : item.customerName || "",
          strNoteStatus: item.strNoteStatus?.trim()
            ? item.strNoteStatus
            : item.strReceptionStatus || "",
          departmentName: item.departmentName?.trim()
            ? item.departmentName
            : item.strProcessingDepartment || "",
          strNatureOfRequest: item.strNatureOfRequest?.trim()
            ? item.strNatureOfRequest
            : item.strNatureofNote || "",
          strNoteTo: item.strNoteTo?.trim()
            ? item.strNoteTo
            : item.sanctioningAuthority || "",
        }));

        setApiData(normalizedData);
        setFilteredData(normalizedData);
        setData(normalizedData);
        setDataResult(process(normalizedData, dataState));
        setDataState({ ...dataState, total: normalizedData.length });
      } else {
        console.error("Error fetching data: Invalid API response", apiData);
        setFilteredData([]);
        setDataResult(process([], dataState));
      }
    } else {
      console.error("Error fetching data:", response.statusText);
    }
  } catch (error) {
    console.error("Error fetching data:", error);
  } finally {
    setIsLoading(false);
  }
};

  const onFilterChange = (ev) => {
    let value = ev.value;

    setFilterValue(value);

    if (!value) {
      // If no filter value, reset to the original data
      setFilteredData(apiData);
      setData(apiData);
    } else {
      let newData = apiData.filter((item) => {
        for (const property in item) {
          if (
            item[property] &&
            item[property].toString &&
            item[property]
              .toString()
              .toLowerCase()
              .includes(value.toLowerCase())
          ) {
            return true;
          }
          if (
            item[property] &&
            item[property].toLocaleDateString &&
            item[property].toLocaleDateString().includes(value)
          ) {
            return true;
          }
        }
        return false;
      });

      setFilteredData(newData);
      setData(newData);

      // Process grouped data if needed
      const newDataResult = processWithGroups(newData, dataState);
      setDataResult(newDataResult);

      // Reset pagination and update total
      const clearedPagerDataState = {
        ...dataState,
        take: 10,
        skip: 0,
      };
      const processedData = process(newData, clearedPagerDataState);
      setDataResult(processedData);
      setDataState({ ...clearedPagerDataState, total: newData.length });
    }
  };

  const [resultState] = React.useState(
    processWithGroups(
      apiData.map((item) => ({
        ...item,
      })),
      initialDataState
    )
  );
  const dataStateChange = (event) => {
    // if (event.dataState && filteredData) {
    setDataResult(process(filteredData, event.dataState));
    setDataState(event.dataState);
    // }
  };
  const onExpandChange = React.useCallback(
    (event) => {
      const newData = [...dataResult.data];
      const item = event.dataItem;
      if (item.groupId) {
        const targetGroup = newData.find((d) => d.groupId === item.groupId);
        if (targetGroup) {
          targetGroup.expanded = event.value;
          setDataResult({
            ...dataResult,
            data: newData,
          });
        }
      } else {
        item.expanded = event.value;
        setDataResult({
          ...dataResult,
          data: newData,
        });
      }
    },
    [dataResult]
  );

    // DELETE FUNCTIONALITY - Similar to eReception
  const onRenderCellDelete = (props) => {
    return (
      <td>
        <Button
          icon="delete"
          onClick={() => handleOpenDialog(props.dataItem)}
          title="Delete Draft"
          className="k-button k-button-md k-rounded-md k-button-solid k-button-solid-base"
        />
      </td>
    );
  };

  const handleOpenDialog = (dataItem) => {
    setDataItemToDelete(dataItem);
        console.log("dataItem to delete:", dataItem);
    setShowDeleteAlert(true);
  };

  const handleConfirmDelete = () => {
    handleDelete(dataItemToDelete);
    setShowDeleteAlert(false);
  };

  const handleDelete = async (item) => {
    const accessToken = await getAccessToken({ ...loginRequest, account }, instance);

    let params = {
      noteId: item.noteId,
    }

    try {
      const response = await fetch(`${API_BASE_URL}${API_ENDPOINTS.eNote_DeleteDraft}`, {
        method: "POST",
        headers: { ...API_COMMON_HEADERS, Authorization: `Bearer ${accessToken}` },
        body: JSON.stringify(params),
      });

      if (response.ok) {
        const updatedData = apiData.filter(
          (dataItem) => dataItem.noteId !== item.noteId
        );

        setApiData(updatedData);
        setFilteredData(updatedData);
        setData(updatedData);
        setDataResult(process(updatedData, dataState));
        setDataState((prevState) => ({
          ...prevState,
          total: updatedData.length,
        }));

        setAlertMsg("Draft note has been deleted successfully");
        setShowNotification(true);
      }
      else {
        console.error("Error deleting note:", response.statusText);
        setAlertMsg("Failed to delete the draft note. Please try again.");
        setShowNotification(true);
      }
    } catch (error) {
      console.error("Error deleting note:", error);
      setAlertMsg("An error occurred while deleting the draft note.");
      setShowNotification(true);
    }
  };


  const exportCSVHeader = (id) => {
    let hdrColumn;

    switch (id) {
      case "All Approved":
        hdrColumn = [
          { key: "noteNumber", label: "Note#" },
          { key: "createdByName", label: "Requester" },
          { key: "departmentName", label: "Department" },
          { key: "subject", label: "Subject" },
          { key: "finalApprover", label: "Final Approver" },
          { key: "createdDate", label: "Created Date" },
          { key: "modifiedDate", label: "Modified Date" },
        ];
        break;

      case "Noted Notes":
        hdrColumn = [
          { key: "noteNumber", label: "Note#" },
          { key: "createdByName", label: "Requester" },
          { key: "departmentName", label: "Department" },
          { key: "subject", label: "Subject" },
          { key: "createdDate", label: "Created Date" },
          { key: "modifiedDate", label: "Modified Date" },
        ];
        break;
      case "All Requests":
        hdrColumn = [
          { key: "noteNumber", label: "Note Number" },
          { key: "createdByName", label: "Requester" },
          { key: "departmentName", label: "Department" },
          { key: "strNoteTo", label: "Note To" },
          { key: "subject", label: "Subject/CustomerName" },
          { key: "currentActionerName", label: "Current Approver" },
          { key: "lastActioner", label: "Last Approver" },
          { key: "strNoteStatus", label: "Status" },
          { key: "modifiedDate", label: "Modified Date" },
          { key: "createdDate", label: "Created Date" },
        ];
        break;
      default:
        hdrColumn = [
          { key: "noteNumber", label: "Note Number" },
          { key: "createdByName", label: "Requester" },
          { key: "departmentName", label: "Department" },
          { key: "strNoteTo", label: "Note To" },
          { key: "subject", label: "Subject" },
          { key: "currentActionerName", label: "Current Approver" },
          { key: "finalApprover", label: "Final Approver" },
          { key: "modifiedDate", label: "Modified Date" },
          { key: "createdDate", label: "Created Date" },
        ];

        // Add "Status" column for "In Progress" view
        if (currentHeading === "In Progress") {
          hdrColumn.splice(5, 0, { key: "strNoteStatus", label: "Status" });
        }
    }
    return hdrColumn;
  };

  const renderColumnsWithData = (data) => {
    if (!data || data.length === 0) {
      return null;
    }

    let columnsConfig = [];

    switch (id) {
     case "All Approved":
   columnsConfig = [
    { field: "noteNumber", title: "Note#", width: 150 },
    { field: "createdByName", title: "Requester", width: 170 },
    { field: "departmentName", title: "Department", width: 180 },
    {
      field: "strNatureOfRequest",
      title: "Nature of Request",
      width: 180,
    },
    {
      field: "strNoteTo",
      title: "Note To",
      width: 170,
    },
    { field: "subject", title: "Subject/CustomerName", width: 330 },
    { field: "finalApprover", title: "Final Approver", width: 180 },
    { field: "modifiedDate", title: "Modified Date", width: 150 },
    { field: "createdDate", title: "Created Date", width: 150 },
  ];
  break;
     case "Noted Notes":
  columnsConfig = [
    { field: "noteNumber", title: "Note#", width: 220 },
    { field: "createdByName", title: "Requester", width: 180 },
    { field: "departmentName", title: "Department", width: 240 },
    { field: "subject", title: "Subject", width: 380 },
    { field: "modifiedDate", title: "Modified Date", width: 220 },
    { field: "createdDate", title: "Created Date", width: 220 },
  ];
  break;

      case "In Progress":
  columnsConfig = [
    { field: "noteNumber", title: "Note#", width: 150 },
    { field: "createdByName", title: "Requester", width: 170 },
    { field: "departmentName", title: "Department", width: 180 },
    {
      field: "strNatureOfRequest",
      title: "Nature of Request",
      width: 180,
    },
    {
      field: "strNoteTo",
      title: "Note To",
      width: 170,
    },
    { field: "subject", title: "Subject/CustomerName", width: 320 },
    {
      field: "currentActionerName",
      title: "Current Approver",
      width: 180,
    },
    { field: "finalApprover", title: "Final Approver", width: 170 },
    { field: "strNoteStatus", title: "Status", width: 130 },
    { field: "modifiedDate", title: "Modified Date", width: 150 },
    { field: "createdDate", title: "Created Date", width: 150 },
  ];
  break;
case "All Requests":
  columnsConfig = [
    { field: "noteNumber", title: "Note Number", width: 150 },
    { field: "createdByName", title: "Requester", width: 170 },
    { field: "departmentName", title: "Department", width: 200 },
    { field: "strNoteTo", title: "Note To", width: 180 },
    { field: "strNatureOfRequest", title: "Nature of Request", width: 180 },
    { field: "subject", title: "Subject/CustomerName", width: 300 },
    { field: "currentActionerName", title: "Current Approver", width: 180 },
    { field: "strNoteStatus", title: "Status", width: 130 },
    { field: "lastActioner", title: "Last Approver", width: 150 },
    { field: "modifiedDate", title: "Modified Date", width: 150 },
    { field: "createdDate", title: "Created Date", width: 150 },
  ];

        break;
    default:
  columnsConfig = [
    { field: "noteNumber", title: "Note Number", width: 160 },
    { field: "createdByName", title: "Requester", width: 160 },
    { field: "departmentName", title: "Department", width: 200 },
    {
      field: "strNoteTo",
      title: "Note To",
      width: 170,
    },
    { field: "subject", title: "Subject/CustomerName", width: 340 },
    {
      field: "currentActionerName",
      title: "Current Approver",
      width: 180,
    },
    { field: "finalApprover", title: "Final Approver", width: 180 },
    { field: "modifiedDate", title: "Modified Date", width: 150 },
    { field: "createdDate", title: "Created Date", width: 150 },
  ];
    }

    return columnsConfig.map((column) => (
      
      <Column
        key={column.field}
        field={column.field}
        title={column.title}
        width={column.width || undefined}
        cell={(props) => {
          const noteFor = props.dataItem["noteFor"];
          console.log(noteFor, "noteFor");
          const status = props.dataItem["status"];
          const createdBy = props.dataItem["createdBy"];
          const noteId = props.dataItem["noteId"];
          const ereceptionId = props.dataItem["ereceptionId"];

          const isDraftOrReturned =
            (status ===
              enumObj.NoteStatus.find((x) => x.dValue === "Draft")?.id ||
              status ===
                enumObj.NoteStatus.find((x) => x.dValue === "Returned")?.id ||
              status ===
                enumObj.NoteStatus.find((x) => x.dValue === "Called Back")
                  ?.id) &&
            createdBy === accounts[0].username;
          console.log("isDraftOrReturned:", isDraftOrReturned);
          console.log("enumObj:", enumObj);
          console.log("props.dataItem:", props.dataItem);

          let linkPath = "";

          switch (noteFor) {
            case 8:
              const basePath = isDraftOrReturned
                ? "eReceptionform"
                : "eReceptionviewform";

              linkPath = `${IB_eReceptionDomain_URL}/${basePath}/${ereceptionId}?from=enote`;
              break;
            case 1:
              linkPath = isDraftOrReturned
                ? `/enoteform/${noteId}`
                : `/enoteviewform/${noteId}`;
              break;
            // case 3:
            //   linkPath = isDraftOrReturned ? `/boardnoteform/${noteId}` : `/boardnoteviewform/${noteId}`;
            //   break;
            default:
              linkPath = "#";
          }

          const linkProps =
            column.field === "noteNumber"
              ? { style: { color: "red" }, className: "_note_number", target: "_blank" }
              : { style: { color: "#424242" } };

          let cellValue;

          if (column.title.includes("Date")) {
            try {
              const rawDate = props.dataItem[column.field];
              cellValue = rawDate
                ? new DateObject(new Date(rawDate)).format(
                    "DD-MMM-YYYY hh:mm A"
                  )
                : "";
            } catch (e) {
              cellValue = props.dataItem[column.field];
            }
          } else {
            if (
              column.field === "strNoteStatus" &&
              ["All Requests", "In Progress"].includes(id)
            ) {
              cellValue = props.dataItem["strNoteStatus"]?.trim()
                ? props.dataItem["strNoteStatus"]
                : props.dataItem["strReceptionStatus"] || "";
            } else if (column.field === "departmentName") {
              if (
                noteFor ===
                enumObj.noteFor.find((x) => x.dValue === "Reception")?.id
              ) {
                cellValue = props.dataItem["strProcessingDepartment"];
              } else {
                cellValue = props.dataItem["departmentName"];
              }
            } else if (column.field === "subject") {
              if (
                noteFor ===
                enumObj.noteFor.find((x) => x.dValue === "Reception")?.id
              ) {
                cellValue = props.dataItem["customerName"];
              } else {
                cellValue = props.dataItem["subject"];
              }
            } else if (column.field === "strNatureOfRequest") {
              if (
                noteFor ===
                enumObj.noteFor.find((x) => x.dValue === "Reception")?.id
              ) {
                cellValue = props.dataItem["strNatureOfRequest"];
              } else {
                cellValue = props.dataItem["strNatureofNote"];
              }
            } else {
              cellValue = props.dataItem[column.field];
            }
          }
          if (column.field === "strNoteTo") {
            if (
              noteFor ===
              enumObj.noteFor.find((x) => x.dValue === "Reception")?.id
            ) {
              cellValue = props.dataItem["sanctioningAuthority"];
            } else {
              cellValue = props.dataItem["strNoteTo"];
            }
          }
          if (column.title === "Amount") {
            try {
              cellValue = formatINRWithCrSuffix(cellValue);
            } catch (e) {
              // fallback
            }
          }

          // ✅ Fallback to 'NA' if value is null, undefined, or empty string
          const finalCellValue =
            cellValue !== null && cellValue !== undefined && cellValue !== ""
              ? cellValue
              : "";

          return (
            <td>
              <Link to={linkPath} {...linkProps}>
                {finalCellValue}
              </Link>
            </td>
          );
        }}
        columnMenu={ColumnMenu}
      />
    ));
  };

  const formatINRWithCrSuffix = (amount) => {
    if (amount) {
      return (
        "₹ " +
        amount.toLocaleString("en-IN", {
          minimumFractionDigits: 2,
          maximumFractionDigits: 2,
        }) +
        " cr"
      );
    }
  };

  return (
    <div>
      <Navbar header="IB Smart Office - eNote" />
      <Sidebar />
      {/* Added top navigation for menu --> 23-09 */}
      <MenuNavContainer isMenuPage />
      <div className="container largeScreen cstGridContainer datagridpage">
        <div className="SectionHeads row mobileSectionHeads">
          {currentHeading}
        </div>
        {/* Add a section for displaying headings and allow the user to switch between them*/}
        {isLoading ? (
          <div>{/* <PageLoader /> */}</div>
        ) : (
          <Grid
            className="cstGridStyles"
            pageable={{ pageSizes: true }}
            data={dataResult}
            sortable={true}
            total={resultState.total}
            onDataStateChange={dataStateChange}
            {...dataState}
            onExpandChange={onExpandChange}
            expandField="expanded"
            dataItemKey={DATA_ITEM_KEY}
            size={"small"}
            resizable={true}
            style={{ height: "auto" }}
            scrollable={"scrollable"}
          >
            <GridToolbar>
              <div className="searchValidationCss">
                <Input
                  value={filterValue}
                  onChange={onFilterChange}
                  className="searchCSS"
                  placeholder="Search in all columns..."
                />
                {filterErrorMessage && (
                  <div className="inCorrectFileError">{filterErrorMessage}</div>
                )}
              </div>

              <div className="export-btns-container">
                <Button className="_exportdiv">
                  <CSVLink
                    filename={`eNote-${selectedView.replace(
                      / /g,
                      ""
                    )}${new DateObject(new Date()).format("DDMMYYYYhhmmss")}`}
                    data={filteredData.map((x) => ({
                      ...x,
                      modifiedDate: new DateObject(
                        new Date(x.modifiedDate)
                      ).format("DD-MMM-YYYY hh:mm A"), // Removed seconds hand
                      createdDate: new DateObject(
                        new Date(x.createdDate)
                      ).format("DD-MMM-YYYY hh:mm A"),
                      // subject: x.subject?.replace(/"/g, '""'), // Escape double quotes
                       strNoteTo:
                       x.noteFor ===
                            enumObj.noteFor.find(
                              (y) => y.dValue === "Reception"
                            )?.id
                          ? x.sanctioningAuthority
                          : x.strNoteTo,

                      subject:
                        x.noteFor ===
                        enumObj.noteFor.find((y) => y.dValue === "Reception")
                          ?.id
                          ? x.customerName
                          : x.subject
                          ? x.subject.replace(/"/g, '""')
                          : "",

                      departmentName:
                        x.noteFor ===
                        enumObj.noteFor.find((y) => y.dValue === "Reception")
                          ?.id
                          ? x.strProcessingDepartment
                          : x.departmentName,

                      strNoteStatus: x.strNoteStatus?.trim()
                        ? x.strNoteStatus
                        : x.strReceptionStatus || "",

                      strNatureOfRequest:
                        x.noteFor ===
                        enumObj.noteFor.find((y) => y.dValue === "Reception")
                          ?.id
                          ? x.strNatureOfRequest
                          : x.strNatureofNote,
                    }))}
                    headers={exportCSVHeader(currentHeading)}
                  >
                    Export CSV
                  </CSVLink>
                </Button>
              </div>
            </GridToolbar>

            {renderColumnsWithData(dataResult)}
            {/* Add Delete Action Column for Draft Requests */}
            {currentHeading === "Draft Requests" && (
              <Column
                field={SELECTED_FIELD}
                width="100px"
                headerClassName="header-checkbox"
                cell={(props) => onRenderCellDelete(props)}
              />
            )}
          </Grid>
        )}
      </div>
      <div className="pgFooterContainer">
        <Footer />
      </div>

        {/* Delete Confirmation Dialog */}
      {showDeleteAlert && (
        <Dialog title="Delete Draft" onClose={() => setShowDeleteAlert(false)} width="40%">
          <div className="dialogAlignment" style={{textAlign: "center"}}><p>Are you sure you want to delete request ?</p></div>
          <DialogActionsBar>
            <Button className="formBtnColor" onClick={handleConfirmDelete}>
              <span className="k-icon k-font-icon  k-i-launch cursor allIconsforPrimary-btn"></span>
              <span>Yes</span>
            </Button>
            <Button onClick={() => setShowDeleteAlert(false)}>
              <span className="k-icon k-font-icon  k-i-close-circle cursor allIconsforPrimary-btn"></span>
              <span>No</span>
            </Button>
          </DialogActionsBar>
        </Dialog>
      )}
      {showNotification && (
        <Dialog title="Delete" onClose={() => setShowNotification(false)}>
          <div className="dialogTextAlignment"><p>{alertMsg}</p></div>
          <DialogActionsBar>
            <Button
              className="formBtnColor notifyDailogOkBtn k-button k-button-md k-rounded-md k-button-solid k-button-solid-base"
              onClick={() => setShowNotification(false)}
            >
              <span className="k-icon k-font-icon  k-i-launch cursor allIconsforPrimary-btn"></span>
              <span>OK</span>
            </Button>
          </DialogActionsBar>
        </Dialog>
      )}
    </div>
  );
};

export default Views;
