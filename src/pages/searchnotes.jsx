import React, { useState, useEffect } from "react";
// Kendo components
import { getter } from "@progress/kendo-react-common";
import { process } from "@progress/kendo-data-query";
import { Input } from "@progress/kendo-react-inputs";
import { Button } from "@progress/kendo-react-buttons";
import {
  Grid,
  GridColumn as Column,
  GridToolbar,
} from "@progress/kendo-react-grid";
import {
  setGroupIds,
  setExpandedState,
} from "@progress/kendo-react-data-tools";
import { DropDownList } from "@progress/kendo-react-dropdowns";
import { Form, FormElement, FieldWrapper } from "@progress/kendo-react-form";
import { Label } from "@progress/kendo-react-labels";
import { DatePicker } from "@progress/kendo-react-dateinputs";
import { Dialog, DialogActionsBar } from "@progress/kendo-react-dialogs";
import { MultiColumnComboBox } from "@progress/kendo-react-dropdowns";
// import kendo icons
import { SvgIcon } from "@progress/kendo-react-common";
import { infoCircleIcon } from "@progress/kendo-svg-icons";
import { xIcon } from "@progress/kendo-svg-icons";
// import external components
import { Link } from "react-router-dom";
import { CSVLink } from "react-csv";
import DateObject from "react-date-object";
import { useMsal, useAccount } from "@azure/msal-react";
// import internal components
import Navbar from "../components/navbar.jsx";
import { Sidebar } from "../components/sidebar.jsx";
import { MenuNavContainer } from "../components/menu";
import Footer from "../components/footer.jsx";
import {
  API_BASE_URL,
  API_ENDPOINTS,
  loginRequest,
  API_COMMON_HEADERS,
  IB_eReceptionDomain_URL,
} from "../config";
import { getAccessToken } from "../App";
import { ColumnMenu } from "./custom-cells.jsx";
import "../styles/searchField.css";


// import style css
import "../styles/datagridpage.css";
import "../styles/forms.css";

const CustomDialogTitleBar = () => {
  return (
    <div className="custom-title">
      <SvgIcon icon={infoCircleIcon} /> Alert!
    </div>
  );
};

//Added this function due to loading impact users will be filtered on search box enter - 25/03
const orgUsersPplPickerColumnMapping = [
  {
    field: "displayName",
    header: "Person Name",
    width: "200px",
  },
  {
    field: "department",
    header: "Department",
    width: "180px",
  },
  {
    field: "jobTitle",
    header: "Designation",
    width: "180px",
  },
  // Change - Adding SR No - 05/04 - Venkat
  {
    field: "srNo",
    header: "SR No",
    width: "120px",
  },
];
// mobile view responsive
const mobileColumns = [
  {
    field: "displayName",
    header: "Person Name",
    width: "100px",
  },
  {
    field: "department",
    header: "Department",
    width: "180px",
  },
  {
    field: "jobTitle",
    header: "Designation",
    width: "100px",
  },
  // Change - Adding SR No - 05/04 - Venkat
  {
    field: "srNo",
    header: "SR No",
    width: "70px",
  },
];

const DATA_ITEM_KEY = "id";
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

export const SearchNotes = () => {
  const valueRender = (element, value, fieldName) => {
    const clearValue = (e) => {
      e.stopPropagation();
      e.preventDefault();
      if (fieldName === "NoteType") {
        setSearchNoteType("");
      }
      if (fieldName === "Status") {
        setSearchNoteStatus("");
      }
      if (fieldName === "FinType") {
        setSearchFinType("");
      }
      if (fieldName === "FY") {
        setSelectedYear("");
      }
      if (fieldName === "Status") {
        setCombinedStatusValue("")
      }
      if (fieldName === "NoteType") {
        setSearchNoteType("")
      }
      if (fieldName === "Branch") {
        setBranchName("");
      }
      if (fieldName === "typeOfConnection") {
        setTypeOfConnectionValue("");
      }
      if (fieldName === "natureOfRequest") {
        setNatureOfRequestValue("");
      }
      if (fieldName === "Committee Name") {
        setSelectedCommitteeValue("");
      }
    };
    if (!value) {
      return element;
    }
    const children = [
      <span key={1} className="_valueRender">
        {element.props.children}
      </span>,
      <SvgIcon icon={xIcon} onClick={clearValue} />,
    ];
    return React.cloneElement(
      element,
      {
        ...element.props,
      },
      children
    );
  };

  const idGetter = getter("id");
  const [filterValue, setFilterValue] = useState("");
  const [filteredData, setFilteredData] = useState([]);
  const [currentSelectedState] = useState({});
  const [dataState, setDataState] = useState(initialDataState);
  const [dataResult, setDataResult] = useState(
    process(filteredData, dataState)
  );
  const [apiData, setApiData] = useState([]);
  const [notestatus, setNoteStatus] = useState([]);
  const [searchNotestatus, setSearchNoteStatus] = useState("");
  const [notetype, setNoteTypeData] = useState([]);
  const [fintype, setFinType] = useState([]);
  const [searchTextNote, setSearchTextNote] = useState("");
  const [searchDepartment, setSearchDepartment] = useState("");
  const [selectedFromDate, setSelectedFromDate] = useState(null);
  const [selectedToDate, setSelectedToDate] = useState(null);
  const [selectedYear, setSelectedYear] = useState("");
  const [searchText, setSearchText] = useState("");
  const [searchSubject, setSearchSubject] = useState("");
  const [searchFinType, setSearchFinType] = useState("");
  const [searchNoteType, setSearchNoteType] = useState("");
  const [searchRequester, setSearchRequester] = useState("");
  const [searchApprover, setSearchApprover] = useState("");
  const [showNotification, setShowNotification] = useState(false);
  const [showNotificationMsg, setShowNotificationMsg] = useState("");
  const [isLoading, setIsLoading] = useState(false);
  const [enumsObj, setEnumsObj] = useState(false);
  const { accounts, instance } = useMsal();
  const account = useAccount(accounts[0] || {});
  const [departmentList, setDepartmentList] = useState([]);
  const [searchTextCustomerName, setSearchTextCustomerName] = useState("");
  const [branchName, setBranchName] = useState("");
  const [typeOfConnectionValue, setTypeOfConnectionValue] = useState("");
  const [natureOfRequestValue, setNatureOfRequestValue] = useState("");
  const [natureOfRequest, setNatureOfRequest] = useState([]);
  const [selectedLastRenewalDate, setSelectedLastRenewalDate] = useState(null);
  const [selectedRenewalDueDate, setSelectedRenewalDueDate] = useState(null);
  const [selectedCommitteeDate, setSelectedCommitteeDate] = useState(null);
  const [committeeDropdownData, setCommitteeDropdownData] = useState([]);
  const [selectedCommitteeValue, setSelectedCommitteeValue] = useState("");
  const [statusEreception, setStatusEreception] = useState([]);
  const [processingDepartment, setProcessingDepartment] = useState([]);
  const [typeOfConnection, setTypeOfConnection] = useState([]);
  const [selectedProcessingDept, setSelectedProcessingDept] = useState("");
  const [combinedStatusValue, setCombinedStatusValue] = useState("");
  const [enumsObjValue, setEnumsObjValue] = useState(false);
  const isMobile = window.innerWidth <= 768;
  const [gridKey, setGridKey] = useState(1);
  // Special chars validation --> (08-10)
  const [searchErrorMessage, setSearchErrorMessage] = useState("");
  const [searchInputBorderColor, setSearchInputBorderColor] = useState(
    "rgba(0, 0, 0, 0.12)"
  );
  const [searchInputErrorMessage, setSearchInputErrorMessage] = useState("");
  const [subjectInputErrorMessage, setSubjectInputErrorMessage] = useState("");
  const [subjectBorderColor, setSubjectBorderColor] = useState(
    "rgba(0, 0, 0, 0.12)"
  );
  const [filterErrorMessage, setFilterErrorMessage] = useState("");
  /* Search Requester and Approver */
  const [members, setMembers] = useState([]);
  const [isSuperAdmin, setIsSuperAdmin] = useState(false);
  const [userDepartment, setUserDepertment] = useState("");

  // Add filteredNatureOfRequest and showRenewalFields state
  const [filteredNatureOfRequest, setFilteredNatureOfRequest] = useState([]);
  useEffect(() => {
    if (typeOfConnectionValue === "New Customer") {
      setFilteredNatureOfRequest(natureOfRequest.filter(x => ["Fresh", "NBG", "BidQuote"].includes(x.dValue)));
    } else if (typeOfConnectionValue === "Existing Customer") {
      setFilteredNatureOfRequest(natureOfRequest.filter(x => !["BidQuote", "Fresh"].includes(x.dValue)));
    } else {
      setFilteredNatureOfRequest(natureOfRequest);
    }
  }, [typeOfConnectionValue, natureOfRequest]);
  const showRenewalFields = ["Review & Renewal", "Enhancement"].includes(natureOfRequestValue);

   const [isAllDataChecked, setIsAllDataChecked] = useState(false);

  // const delay = 100
  useEffect(() => {
    const fetchData = async () => {
      setIsLoading(true);
      try {
        const accessToken = await getAccessToken(
          { ...loginRequest, account },
          instance
        );

        // Fetch departments list
        const departmentResponse = await fetch(
          `${API_BASE_URL}${API_ENDPOINTS.GET_Departments_List}`,
          {
            method: "GET",
            headers: {
              ...API_COMMON_HEADERS,
              Authorization: `Bearer ${accessToken}`,
            },
          }
        );
        const departmentDetailsList = await departmentResponse.json();
        setDepartmentList(departmentDetailsList);

        // Fetch super admin status
        const deptAdminResponse = await fetch(
          `${API_BASE_URL}${API_ENDPOINTS.Find_Super_Dept_Admin(
            accounts[0]?.username
          )}`,
          {
            method: "GET",
            headers: {
              ...API_COMMON_HEADERS,
              Authorization: `Bearer ${accessToken}`,
            },
          }
        );
        const deptAdminData = await deptAdminResponse.json();
        setIsSuperAdmin(deptAdminData.isSuperAdmin);

        if (!deptAdminData.isSuperAdmin) {
          // Find the department for non-super admins
          const userDepartment = departmentDetailsList.find(
            (dept) => dept.departmentId === deptAdminData.departmentId
          );

          if (userDepartment) {
            setSearchDepartment(userDepartment.departmentName); // Use departmentName
          } else {
            console.warn(
              "No matching department found for departmentId:",
              deptAdminData.departmentId
            );
            setSearchDepartment(""); // Fallback to an empty string
          }
        }

        const getDepartmentList = async () => {
          const accessToken = await getAccessToken({ ...loginRequest, account }, instance);

          const obj = await fetch(`${API_BASE_URL}${API_ENDPOINTS.GET_Departments_List}`
            , {
              method: "GET",
              headers: { ...API_COMMON_HEADERS, Authorization: `Bearer ${accessToken}` }
            }
          );
          const departmentDetailsList = await obj.json();
          setDepartmentList(departmentDetailsList);

        }
        const ereceptiondropdowns = await fetch(`${API_BASE_URL}${API_ENDPOINTS.GET_DROPDOWNDATA_ereception}`, {
          method: "GET",
          headers: { ...API_COMMON_HEADERS, Authorization: `Bearer ${accessToken}` }
        });

        const dropdownslistDetails = await ereceptiondropdowns.json();
        getDepartmentList();
        setEnumsObjValue(dropdownslistDetails);
        setIsLoading(false);
        setProcessingDepartment(dropdownslistDetails.ProcessingDepartment);
        setNatureOfRequest(dropdownslistDetails.NatureOfRequest);
        setTypeOfConnection(dropdownslistDetails.TypeOfConnection);
        setSelectedProcessingDept(dropdownslistDetails.SelectedProcessingDept);
        setSelectedLastRenewalDate(dropdownslistDetails.SelectedLastRenewalDate);
        setSelectedRenewalDueDate(dropdownslistDetails.SelectedRenewalDueDate);
        setSelectedCommitteeDate(dropdownslistDetails.SelectedCommitteeDate);
        setStatusEreception(dropdownslistDetails.ReceptionSearchStatus);
        setCommitteeDropdownData(dropdownslistDetails.committeeName);

        //get enum objects
        const dropdowns = await fetch(`${API_BASE_URL}${API_ENDPOINTS.GET_DROPDOWNDATA}`,
          {
            method: "GET",
            headers: {
              ...API_COMMON_HEADERS,
              Authorization: `Bearer ${accessToken}`,
            },
          }
        );

        const dropdownslist = await dropdowns.json();
        setNoteTypeData(dropdownslist.SearchNoteType);
        setFinType(dropdownslist.SearchFinancial);
        setNoteStatus(dropdownslist.SearchStatus);
        setEnumsObj(dropdownslist);
        setIsLoading(false);
      } catch (error) {
        console.error("Error fetching data:", error);
      }
    };
    fetchData();
    findDeptSuperAdmin();
    getUserDepartment();
  }, []);



  const findDeptSuperAdmin = async () => {
    const accessToken = await getAccessToken(
      { ...loginRequest, account },
      instance
    );

    try {
      const response = await fetch(
        `${API_BASE_URL}${API_ENDPOINTS.Find_Super_Dept_Admin(
          accounts[0]?.username
        )}`,
        {
          method: "GET",
          headers: {
            ...API_COMMON_HEADERS,
            Authorization: `Bearer ${accessToken}`,
          },
        }
      );

      const data = await response.json();
      setIsSuperAdmin(data.isSuperAdmin); // Update state
    } catch (error) {
      console.error("Error fetching department admin:", error);
    }
  };

  const onFilterChange = (ev) => {
    let value = ev.value;
    setFilterValue(value);

    if (!value) {
      setFilteredData(apiData);
      setDataResult(
        process(apiData, (dataState) => ({
          ...dataState,
          total: apiData.length,
        }))
      );
    } else {
      let newData = apiData.filter((item) => {
        for (const property in item) {
          if (
            item[property] &&
            item[property].toString &&
            item[property]
              .toString()
              .toLocaleLowerCase()
              .includes(value.toLocaleLowerCase())
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
        selected: currentSelectedState[idGetter(item)],
      })),
      initialDataState
    )
  );
  const dataStateChange = (event) => {
    setDataResult(process(filteredData, event.dataState));
    setDataState(event.dataState);
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

  // Handle selected Value change
  const setSelectedValue = (data) => {
    let newData = data.map((item) => {
      if (item.items) {
        return {
          ...item,
          items: setSelectedValue(item.items),
        };
      } else {
        return {
          ...item,
          selected: currentSelectedState[idGetter(item)],
        };
      }
    });
    return newData;
  };
  // const newData =
  setExpandedState({
    data: setSelectedValue(resultState.data),
    collapsedIds: [],
  });

  // Handle export CSV headers
  const exportCSVHeader = () => {
    return [
      { key: "noteNumber", label: "Note#" },
      { key: "createdBy", label: "Requester" },
      { key: "departmentName", label: "Department" },
      { key: "strNoteTo", label: "Note To" },
      { key: "subject", label: "Customer Name" },
      { key: "strNoteStatus", label: "Status" },
      { key: "currentActioner", label: "Current Approver" },
      { key: "modifiedBy", label: "Last Approver" },
      { key: "modifiedDate", label: "Modified Date" },
      { key: "createdDate", label: "Created Date" },
    ];
  };

  // Handle  render columns with data
  const renderColumnsWithData = (data) => {
    if (!data || data.length === 0) {
      return null;
    }

    const columnsConfig = [
      { field: "noteNumber", title: "Note#", width: 150 },
      { field: "createdBy", title: "Requester", width: 200 },
      { field: "departmentName", title: "Department", width: 200 },
      { field: "strNoteTo", title: "Note To", width: 200 },
      { field: "subject", title: "Subject/Customer Name", width: 350 },
      { field: "strNoteStatus", title: "Status", width: 150 },
      { field: "currentActioner", title: "Current Approver", width: 200 },
      { field: "modifiedBy", title: "Last Approver", width: 200 },
      { field: "modifiedDate", title: "Modified Date", width: 180 },
      { field: "createdDate", title: "Created Date", width: 180 },
    ];

    return columnsConfig.map((column) => (
      <Column
        key={column.field}
        field={column.field}
        title={column.title}
        width={column.width || undefined}
        cell={(props) => {
          const noteFor = props.dataItem["noteForId"];
          const status = props.dataItem["status"];
          const createdBy = props.dataItem["createdBy"];
          const noteId = props.dataItem["noteId"];
          const ereceptionId = props.dataItem["ereceptionId"];

          const isDraftOrReturned =
            (status === enumsObj.NoteStatus.find((x) => x.dValue === "Draft")?.id ||
              status === enumsObj.NoteStatus.find((x) => x.dValue === "Returned")?.id ||
              status === enumsObj.NoteStatus.find((x) => x.dValue === "Called Back")?.id) &&
            createdBy === accounts[0].username;

          let linkPath = "#";

          switch (noteFor) {
            case 8:
              const basePath = isDraftOrReturned
                ? "eReceptionform"
                : "eReceptionviewform";

              linkPath = `${IB_eReceptionDomain_URL}/${basePath}/${ereceptionId}?from=enote`;
              break;
            case 1:
              linkPath = isDraftOrReturned ? `/enoteform/${noteId}` : `/enoteviewform/${noteId}`;
              break;
            default:
              linkPath = "#";
          }

          const linkProps = column.field === "noteNumber" ? {
            style: { color: "red" },
            className: "_note_number",
            target: "_blank",
          }
            : { style: { color: "#424242" }, target: "_blank" };

          let rawValue;

          // Format date fields
          if (column.title.includes("Date")) {
            try {
              const rawDate = props.dataItem[column.field];
              rawValue = rawDate
                ? new DateObject(new Date(rawDate)).format(
                  "DD-MMM-YYYY hh:mm A"
                )
                : "";
            } catch {
              rawValue = props.dataItem[column.field];
            }
          }
          // Special formatting for boardCommitteeName in board context
          else if (column.field === "strNoteTo") {
            if (
              noteFor === enumsObj.noteFor.find((x) => x.dValue === "Reception")?.id
            ) {
              rawValue = props.dataItem["sanctioningAuthority"];
            } else {
              rawValue = props.dataItem["strNoteTo"];
            }
          }
          // Format amount
          else if (column.title === "Amount") {
            try {
              rawValue = formatINRWithCrSuffix(props.dataItem[column.field]);
            } catch {
              rawValue = props.dataItem[column.field];
            }
          }
          // Status logic
          else if (column.field === "strNoteStatus") {
            rawValue = props.dataItem["strNoteStatus"]?.trim()
              ? props.dataItem["strNoteStatus"]
              : props.dataItem["strReceptionStatus"] || "";
          } else if (column.field === "departmentName") {
            if (
              noteFor === enumsObj.noteFor.find((x) => x.dValue === "Reception")?.id
            ) {
              rawValue = props.dataItem["strProcessingDepartment"];
            } else {
              rawValue = props.dataItem["departmentName"];
            }
          } else if (column.field === "subject") {
            if (
              noteFor === enumsObj.noteFor.find((x) => x.dValue === "Reception")?.id
            ) {
              rawValue = props.dataItem["customerName"];
            } else {
              rawValue = props.dataItem["subject"];
            }
          } else if (column.field === "strNatureOfRequest") {
            if (
              noteFor === enumsObj.noteFor.find((x) => x.dValue === "Reception")?.id
            ) {
              rawValue = props.dataItem["strNatureOfRequest"];
            } else {
              rawValue = props.dataItem["strNatureofNote"];
            }
          }
          // Default
          else {
            rawValue = props.dataItem[column.field];
          }

          // ✅ Final fallback to 'NA' if value is null, undefined, or empty string
          const cellValue =
            rawValue !== null && rawValue !== undefined && rawValue !== ""
              ? rawValue
              : "";

          return (
            <td>
              <Link to={linkPath} {...linkProps}>
                {cellValue}
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

  /* To get the financial year viz in 2024-2025 format */
  // const getRecentFinancialYears = () => {
  //   const currentYear = new Date().getFullYear();
  //   const years = [];
  //   for (let i = 0; i < 5; i++) {
  //     const startYear = currentYear - i;
  //     const endYear = startYear + 1; // Increment startYear by 1 to get the endYear
  //     years.push({ text: `${startYear}-${endYear}`, value: startYear });
  //   }
  //   return years;
  // };

  const getRecentFinancialYears = () => {
    const startYear = 2020;
    const currentYear = new Date().getFullYear();
    const years = [];

    for (let year = startYear; year < currentYear; year++) {
      years.push({
        text: `${year}-${year + 1}`,
        value: year
      });
    }

    return years;
  };

  const financialYears = getRecentFinancialYears();
  const combinedStatus = [...notestatus, ...statusEreception];
  const uniqueCombinedStatus = Array.from(
    new Map(combinedStatus.map(item => [item.dValue, item]))
  ).map(([_, item]) => item);
  /* Search Functionality */
  const handleSearch = async () => {
    if (
      searchTextNote === "" &&
      (searchRequester === "" || searchRequester === null) &&
      searchDepartment === "" &&
      userDepartment === " " &&
      searchText === "" &&
      selectedFromDate === null &&
      selectedToDate === null &&
      selectedYear === "" &&
      searchNotestatus === "" &&
      searchSubject === "" &&
      searchFinType === "" &&
      (searchApprover === "" || searchApprover === null) &&
      searchNoteType === "" &&
      selectedCommitteeDate === null &&
      selectedRenewalDueDate === null &&
      selectedLastRenewalDate === null &&
      selectedYear === "" &&
      searchSubject === "" &&
      searchFinType === "" &&
      searchTextCustomerName === "" &&
      natureOfRequestValue === "" &&
      (searchApprover === "" || searchApprover === null) &&
      // Check dropdown fields properly
      (typeOfConnectionValue === "" || typeOfConnectionValue === null || typeOfConnectionValue === undefined) &&
      (selectedProcessingDept === "" || selectedProcessingDept === null || selectedProcessingDept === undefined) &&
      (searchApprover.userPrincipalName === '' || searchApprover.userPrincipalName === null)
    ) {
      setShowNotification(true);
      setShowNotificationMsg(
        "Please fill at least any one of the fields to search."
      );
      return;
    }

    setIsLoading(true);
    try {
      const financialyear = selectedYear ? selectedYear.text : null;

      // Determine the department name based on super admin status
      const departmentName = isSuperAdmin
        ? (searchDepartment ? searchDepartment : null)
        : userDepartment;
      const selectedStatus = combinedStatus.find(
        x => x.dValue === combinedStatusValue
      );
      let statusParam = 0;
      let receptionStatusParam = 0;

      if (
        Array.isArray(enumsObj?.SearchStatus) &&
        enumsObj.SearchStatus.some((x) => x.dValue === selectedStatus?.dValue)
      ) {
        statusParam =
          enumsObj.SearchStatus.find((x) => x.dValue === selectedStatus?.dValue)?.id || 0;
      }

      if (
        Array.isArray(enumsObj?.ReceptionSearchStatus) &&
        enumsObj.ReceptionSearchStatus.some((x) => x.dValue === selectedStatus?.dValue)
      ) {
        receptionStatusParam =
          enumsObj.ReceptionSearchStatus.find((x) => x.dValue === selectedStatus?.dValue)?.id || 0;
      }


      // Create an object with all possible fields
      let params = {
        noteNumber: searchTextNote || null,
        createdBy: searchRequester?.userPrincipalName || null,
        departmentName: departmentName,
        searchKeyword: searchText || null,
        fromDate: selectedFromDate ? new DateObject(new Date(selectedFromDate)).format("DD-MM-YYYY") : null,
        toDate: selectedToDate ? new DateObject(new Date(selectedToDate)).format("DD-MM-YYYY") : null,
        status: statusParam,
        receptionStatus: receptionStatusParam,
        subject: searchSubject || null,
        financialType: searchFinType ? enumsObj.SearchFinancial.find((x) => x.dValue === searchFinType).id : 0,
        fy: financialyear || null,
        approverEmail: searchApprover?.userPrincipalName || null,
        natureofNote: searchNoteType ? enumsObj.SearchNoteType.find((x) => x.dValue === searchNoteType).id : 0,
        loginMailID: accounts[0].username,
        requestNumber: searchTextNote || null,
        customerName: searchTextCustomerName || null,
        processingDepartment: branchName ? enumsObjValue?.ProcessingDepartment.find((x) => x.dValue === branchName)?.id || 0 : 0,
        typeOfConnection: typeOfConnectionValue ? enumsObjValue?.TypeOfConnection.find((x) => x.dValue === typeOfConnectionValue)?.id || 0 : 0,
        approversEmail: searchApprover?.userPrincipalName || null,
        natureOfRequest: natureOfRequestValue ? enumsObjValue?.NatureOfRequest.find((x) => x.dValue === natureOfRequestValue)?.id || 0 : 0,
        lastRenewaDate: selectedLastRenewalDate ? new DateObject(new Date(selectedLastRenewalDate)).format("DD-MM-YYYY") : null,
        renewalDueOnDate: selectedRenewalDueDate ? new DateObject(new Date(selectedRenewalDueDate)).format("DD-MM-YYYY") : null,
        committeeDateDess: selectedCommitteeDate ? new DateObject(new Date(selectedCommitteeDate)).format("DD-MM-YYYY") : null,
        committeeNameId: selectedCommitteeValue ? enumsObjValue?.committeeName.find((x) => x.dValue === selectedCommitteeValue)?.id || 0 : 0,
      IncludeArchivedNote: isAllDataChecked
      };

      // Set default values for financialType, status, and natureofNote if they are 0
      if (!params.financialType) {
        params.financialType = 0;
      }
      if (!params.status) {
        params.status = 0;
      }
      if (!params.natureofNote) {
        params.natureofNote = 0;
      }

      // Remove other fields that are null or empty strings
      params = Object.fromEntries(
        Object.entries(params).filter(
          ([key, value]) =>
            value !== null &&
            value !== "" &&
            (value !== 0 ||
              ["financialType", "status", "natureofNote"].includes(key))
        )
      );

      const accessToken = await getAccessToken({ ...loginRequest, account }, instance);

      await fetch(`${API_BASE_URL}${API_ENDPOINTS.eNote_SearchNotes}`, {
        method: "POST",
        body: JSON.stringify(params),
        headers: {
          ...API_COMMON_HEADERS,
          Authorization: `Bearer ${accessToken}`,
        },
      })
        .then((response) => response.json())
        .then((data) => {
          setIsLoading(false);
          const apiRespData = data;

          setFilteredData(apiRespData);
          setApiData(apiRespData);
          const newDataResult = process(apiRespData, dataState);
          setDataResult(newDataResult);
          setDataState({ ...dataState, total: apiRespData.length });
        });
    } catch (error) {
      console.error("Error fetching data:", error);
      setIsLoading(false);
    }
  };

  const handleClear = () => {
    setGridKey(prev => prev + 1);
         setIsAllDataChecked(false); 
    setFilterValue("");
    setFilteredData([]);
    setApiData([]);

    const clearedState = {
      ...initialDataState,
      filter: { logic: "and", filters: [] }
    };
    setDataState(clearedState);
    setDataResult(process([], clearedState));
    setFilterValue("");
    setFilteredData([]);
    setApiData([]);
    setDataResult(process([], initialDataState)); // Reset data state
    setSearchTextNote("");
    setSearchDepartment("");
    setSearchText("");
    setSelectedFromDate(null);
    setSelectedToDate(null);
    setSearchRequester({ department: "", displayName: "", jobTitle: "", userPrincipalName: "" });
    setSearchSubject("");
    setSelectedYear(null);
    setSearchNoteStatus("");
    setSearchApprover({ department: "", displayName: "", jobTitle: "", userPrincipalName: "", });
    setSearchFinType("");
    setSearchNoteType("");
    setFilterErrorMessage("");
    setSearchErrorMessage("");
    setSearchInputBorderColor("");
    setSearchInputErrorMessage("");
    setSubjectBorderColor("");
    setSubjectInputErrorMessage("");
    setSelectedProcessingDept(null);
    setSelectedLastRenewalDate(null);
    setSelectedRenewalDueDate(null);
    setSelectedCommitteeDate(null);
    setNatureOfRequestValue(null);
    setTypeOfConnectionValue(null);
    setSearchText("");
    setSearchTextCustomerName("");
    setSearchSubject("");
    setSelectedCommitteeValue("");
    setCombinedStatusValue("");
    setBranchName("");
  };

  //filter the users dropdown data
  const filterChange = async (event) => {
    if (event.filter.value.length >= 4) {
      const accessToken = await getAccessToken({ ...loginRequest, account }, instance);

      await fetch(`${API_BASE_URL}${API_ENDPOINTS.Search_UserDetails(event.filter.value)}`, {
        method: "GET",
        headers: {
          ...API_COMMON_HEADERS,
          Authorization: `Bearer ${accessToken}`,
        },
      })
        .then((response) => {
          return response.json();
        })
        .then((data) => {
          const orgUsers = data.map((x) => {
            // Change - Adding SR No - 05/04 - Venkat
            return {
              ...x,
              department: x.department === null ? "NA" : x.department,
              displayName: x.displayName === null ? "NA" : x.displayName,
              jobTitle: x.jobTitle === null ? "NA" : x.jobTitle,
              userPrincipalName: x.userPrincipalName,
            };
          });
          setMembers(orgUsers);
        })
        .catch((err) => {
          setMembers([]);
          console.log(err);
        });
    }
  };

  // Special chars validation --> (08-10)
  const handleInputChange = (e) => {
    const value = e.target.value;
    setSearchTextNote(value);
  };

  const handleInputSearchChange = (e) => {
    const value = e.target.value;
    setSearchText(value);
  };

  const handleInputSubjectChange = (e) => {
    const value = e.target.value;
    setSearchSubject(value);
  };

  const getUserDepartment = async () => {
    try {
      const accessToken = await getAccessToken({ ...loginRequest, account }, instance);

      const obj = await fetch(
        `${API_BASE_URL}${API_ENDPOINTS.GET_UserDetailsByPrincipalName(
          accounts[0].username
        )}`,
        {
          method: "GET",
          headers: {
            ...API_COMMON_HEADERS,
            Authorization: `Bearer ${accessToken}`,
          },
        }
      );
      const departmentDetails = await obj.json();
      // Commentted UserDepartment State in onload
      setUserDepertment(departmentDetails[0].department);
    } catch (error) {
      console.error("Error fetching user details:", error);
    }
  };

  return (
    <div>
      <Navbar />
      <Sidebar />
      {/* Added top navigation for menu --> 23-09 */}
      <MenuNavContainer isMenuPage />
      <div className="container largeScreen datagridpage searchNoteContainer">
        <div className="SectionHeads row mobileSectionHeads">
          Search Parameters
        </div>
        <div className="container largeScreen">
          <Form
            id="searchForm"
            render={() => (
              <FormElement>
                <fieldset className={"k-form-fieldset"}>
                  <div className="SectionRow row">
                    <div className="col-md-3">
                      <FieldWrapper>
                        <div className="k-form-field-wrap">
                          <Label className="k-form-label">Note#:</Label>
                          <Input
                            name="Note#"
                            value={searchTextNote}
                            onChange={handleInputChange}
                          />
                          {searchErrorMessage && (
                            <div className="inCorrectFileError">
                              {searchErrorMessage}
                            </div>
                          )}
                        </div>
                      </FieldWrapper>
                    </div>
                    <div className="col-md-3">
                      <FieldWrapper>
                        <div className="k-form-field-wrap">
                          <Label className="k-form-label">Requester:</Label>
                          <MultiColumnComboBox
                            name="Requester"
                            data={members}
                            filterable={true}
                            className="testt"
                            columns={
                              isMobile
                                ? mobileColumns
                                : orgUsersPplPickerColumnMapping
                            }
                            style={{ width: isMobile ? "100%" : "100%" }} // Adjust width
                            onFilterChange={filterChange}
                            value={searchRequester.displayName}
                            onChange={(e) =>
                              e.value === null
                                ? setSearchRequester({
                                  department: "",
                                  displayName: "",
                                  jobTitle: "",
                                  userPrincipalName: "",
                                })
                                : setSearchRequester(e.value)
                            }
                            placeholder="Search Requester..."
                          />
                          <span className="cstUserSearchMsg">
                            (Please enter minimum 4 characters to find the user)
                          </span>
                        </div>
                      </FieldWrapper>
                    </div>
                    <div className="col-md-3">
                      <FieldWrapper>
                        <div className="k-form-field-wrap">
                          <Label className="k-form-label">Approver:</Label>
                          <MultiColumnComboBox
                            name="Approver"
                            data={members}
                            filterable={true}
                            columns={
                              isMobile
                                ? mobileColumns
                                : orgUsersPplPickerColumnMapping
                            }
                            className="responsive-multicolumn"
                            style={{ width: isMobile ? "100%" : "100%" }} // Adjust width
                            onFilterChange={filterChange}
                            value={searchApprover.displayName}
                            onChange={(e) =>
                              e.value === null
                                ? setSearchApprover({
                                  department: "",
                                  displayName: "",
                                  jobTitle: "",
                                  userPrincipalName: "",
                                })
                                : setSearchApprover(e.value)
                            }
                            placeholder="Search Approver..."
                          />
                          <span className="cstUserSearchMsg">
                            (Please enter minimum 4 characters to find the user)
                          </span>
                        </div>
                      </FieldWrapper>
                    </div>
                    <div className="col-md-3">
                      {!isSuperAdmin ? (
                        <FieldWrapper>
                          <div className="k-form-field-wrap">
                            <Label className="k-form-label">Department:</Label>
                            <Input
                              name="Department"
                              value={userDepartment}
                              readOnly
                            />
                          </div>
                        </FieldWrapper>
                      ) : (
                        <FieldWrapper>
                          <div className="k-form-field-wrap">
                            <Label className="k-form-label">Department:</Label>
                            <DropDownList
                              data={departmentList.map((x) => x.departmentName)}
                              onChange={(e) =>
                                setSearchDepartment(e.target.value)
                              }
                              value={searchDepartment}
                            />
                          </div>
                        </FieldWrapper>
                      )}
                    </div>

                    <div className="col-md-3">
                      <FieldWrapper>
                        <div className="k-form-field-wrap">
                          <Label className="k-form-label">Search Text:</Label>
                          <Input
                            name="search"
                            value={searchText}
                            onChange={handleInputSearchChange}
                            style={{ borderColor: searchInputBorderColor }}
                          />
                          {searchInputErrorMessage && (
                            <div className="inCorrectFileError">
                              {searchInputErrorMessage}
                            </div>
                          )}
                        </div>
                      </FieldWrapper>
                    </div>
                    <div className="col-md-3">
                      <FieldWrapper>
                        <div className="k-form-field-wrap">
                          <Label className="k-form-label">From Date:</Label>
                          <DatePicker
                            max={new Date()}
                            placeholder="From Date..."
                            onChange={(event) =>
                              setSelectedFromDate(event.target.value)
                            }
                            format={"dd-MM-yyyy"}
                            value={selectedFromDate}
                          />
                        </div>
                      </FieldWrapper>
                    </div>
                    <div className="col-md-3">
                      <FieldWrapper>
                        <div className="k-form-field-wrap">
                          <Label className="k-form-label">To Date:</Label>
                          <DatePicker
                            min={new Date(selectedFromDate)}
                            max={new Date()}
                            placeholder="To Date..."
                            onChange={(event) =>
                              setSelectedToDate(event.target.value)
                            }
                            format={"dd-MM-yyyy"}
                            value={selectedToDate}
                          />
                        </div>
                      </FieldWrapper>
                    </div>
                    <div className="col-md-3">
                      <FieldWrapper>
                        <div className="k-form-field-wrap">
                          <Label className="k-form-label">Status:</Label>
                          <DropDownList
                            data={uniqueCombinedStatus.map((x) => x.dValue)}
                            value={combinedStatusValue}
                            onChange={(e) => setCombinedStatusValue(e.target.value)}
                            valueRender={(element) => valueRender(element, combinedStatusValue, "Status")}
                          />
                        </div>
                      </FieldWrapper>
                    </div>
                    <div className="col-md-3">
                      <FieldWrapper>
                        <div className="k-form-field-wrap">
                          <Label className="k-form-label">Subject:</Label>
                          <Input
                            name="Subject"
                            value={searchSubject}
                            onChange={handleInputSubjectChange}
                            style={{ borderColor: subjectBorderColor }}
                          />
                          {subjectInputErrorMessage && (
                            <div className="inCorrectFileError">
                              {subjectInputErrorMessage}
                            </div>
                          )}
                        </div>
                      </FieldWrapper>
                    </div>
                    <div className="col-md-3">
                      <FieldWrapper>
                        <div className="k-form-field-wrap">
                          <Label className="k-form-label">Financial:</Label>
                          <DropDownList
                            data={fintype.map((x) => x.dValue)}
                            value={searchFinType}
                            onChange={(e) => setSearchFinType(e.target.value)}
                            valueRender={(element) =>
                              valueRender(element, searchFinType, "FinType")
                            }
                          />
                        </div>
                      </FieldWrapper>
                    </div>
                    <div className="col-md-3">
                      <FieldWrapper>
                        <div className="k-form-field-wrap">
                          <Label className="k-form-label">FY:</Label>
                          <DropDownList
                            data={financialYears}
                            textField="text"
                            dataItemKey="text"
                            value={selectedYear}
                            onChange={(event) =>
                              setSelectedYear(event.target.value)
                            }
                            valueRender={(element) =>
                              valueRender(element, selectedYear, "FY")
                            }
                          />
                        </div>
                      </FieldWrapper>
                    </div>


                    <div className="col-md-3">
                      <FieldWrapper>
                        <div className="k-form-field-wrap">
                          <Label className="k-form-label">Note Type:</Label>
                          <DropDownList
                            data={notetype.map((x) => x.dValue)}
                            value={searchNoteType}
                            onChange={(e) => setSearchNoteType(e.target.value)}
                            valueRender={(element) =>
                              valueRender(element, searchNoteType, "NoteType")
                            }
                          />
                        </div>
                      </FieldWrapper>
                    </div>
                    <div className="col-md-3" >
                      <FieldWrapper>
                        <div className="k-form-field-wrap">
                          <Label className="k-form-label">Customer Name</Label>
                          <Input
                            spellCheck={true}
                            name="CustomerName"
                            value={searchTextCustomerName}
                            onChange={(e) => setSearchTextCustomerName(e.target.value)}
                          />
                        </div>
                      </FieldWrapper>
                    </div>
                    <div className="col-md-3">
                      <FieldWrapper>
                        <div className="k-form-field-wrap">
                          <Label className="k-form-label">Branch</Label>
                          <DropDownList
                            data={processingDepartment.map((x) => x.dValue)}
                            value={branchName}
                            onChange={(e) => setBranchName(e.target.value)}
                            valueRender={element => valueRender(element, branchName, "Branch")}
                          />
                        </div>
                      </FieldWrapper>
                    </div>
                    <div className="col-md-3" >
                      <FieldWrapper>
                        <div className="k-form-field-wrap">
                          <Label className="k-form-label">Type of Connection</Label>
                          <DropDownList
                            data={typeOfConnection.map((x) => x.dValue)}
                            onChange={(e) => setTypeOfConnectionValue(e.target.value)}
                            value={typeOfConnectionValue}
                            name="typeOfConnection"
                            valueRender={element => valueRender(element, typeOfConnectionValue, "typeOfConnection")}
                          />
                        </div>
                      </FieldWrapper>
                    </div>
                    <div className="col-md-3" >
                      <FieldWrapper>
                        <div className="k-form-field-wrap">
                          <Label className="k-form-label"> Nature of Request:</Label>
                          <DropDownList
                            data={filteredNatureOfRequest.map((x) => x.dValue)}
                            value={natureOfRequestValue}
                            onChange={(e) => setNatureOfRequestValue(e.target.value)}
                            valueRender={(element) => valueRender(element, natureOfRequestValue, "natureOfRequest")}
                          />
                        </div>
                      </FieldWrapper>
                    </div>
                    {/* Last Renewal Date (conditional) */}
                    {showRenewalFields && (
                      <div className="col-md-3" >
                        <FieldWrapper>
                          <div className="k-form-field-wrap">
                            <Label className="k-form-label">Last Renewal Date</Label>
                            <DatePicker
                              max={new Date()}
                              placeholder="Last Renewal Date..."
                              onChange={(e) => setSelectedLastRenewalDate(e.target.value)}
                              format="dd-MM-yyyy"
                              value={selectedLastRenewalDate}
                            />
                          </div>
                        </FieldWrapper>
                      </div>
                    )}
                    {/* Renewal Due On (conditional) */}
                    {showRenewalFields && (
                      <div className="col-md-3" >
                        <FieldWrapper>
                          <div className="k-form-field-wrap">
                            <Label className="k-form-label">Renewal Due On</Label>
                            <DatePicker
                              max={new Date()}
                              placeholder="Renewal Due On..."
                              onChange={(e) => setSelectedRenewalDueDate(e.target.value)}
                              format="dd-MM-yyyy"
                              value={selectedRenewalDueDate}
                            />
                          </div>
                        </FieldWrapper>
                      </div>
                    )}
                    {/* Committee Name */}
                    <div className="col-md-3">
                      <FieldWrapper>
                        <div className="k-form-field-wrap">
                          <Label className="k-form-label">Sanctioning Authorities</Label>
                          <DropDownList
                            data={committeeDropdownData
                              .map((x) => x.dValue)
                              .filter((val) => ["MD", "ED", "CGM", "GM"].includes(val))}
                            value={selectedCommitteeValue}
                            onChange={(e) => setSelectedCommitteeValue(e.target.value)}
                            valueRender={(element) =>
                              valueRender(element, selectedCommitteeValue, "Committee Name")
                            }
                          />
                        </div>
                      </FieldWrapper>
                    </div>

                          <div className="col-md-3">
  <FieldWrapper>
    <div className="k-form-field-wrap">

      {/* Dummy label to maintain alignment */}
      <Label className="k-form-label">&nbsp;</Label>

      <label className="Check-box-field checkbox-align">
        <input
          type="checkbox"
          className="big-checkbox"
          checked={isAllDataChecked}
          onChange={(e) => setIsAllDataChecked(e.target.checked)}
        />
        <span className='ArchiveNote'>Include Archived Notes</span>
      </label>

    </div>
  </FieldWrapper>
</div>

                  </div>
                </fieldset>
                <div className="FormButtonsContainer k-form-buttons">
                  <Button className="searchNoteBtns" onClick={handleSearch}>
                    <span className="k-icon-xs k-icon k-font-icon  k-i-search cursor searchBtnStyles"></span>
                    Search
                  </Button>
                  <Button onClick={handleClear}>
                    <span className="k-icon-sm k-icon k-font-icon  k-i-reset-sm cursor searchBtnStyles"></span>
                    Clear
                  </Button>
                </div>
              </FormElement>
            )}
          />
        </div>
        <div className="SectionHeads row _search_div">Search Results</div>
        {isLoading ? (
          <div>{/* <PageLoader /> */}</div>
        ) : (
          <Grid
            // Added mobile view responsive
            key={gridKey}
            className="cstGridStyles searchMobileView"
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
            
          >
            <GridToolbar>
              <div className="searchValidationCss">
                <Input
                  value={filterValue}
                  onChange={onFilterChange}
                  className="searchCSS"
                  placeholder="Search in all columns..."
                />{" "}
                {filterErrorMessage && (
                  <div className="inCorrectFileError">{filterErrorMessage}</div>
                )}
              </div>
              <div className="export-btns-container">
                <Button className="_exportdiv">
                  <CSVLink
                    filename={`eNote-SearchResult-${new DateObject(
                      new Date()
                    ).format("DDMMYYYYhhmmss")}`}
                    data={filteredData.map((x) => ({
                      ...x,
                      strNoteTo: x.noteFor === enumsObj.noteFor.find((y) => y.dValue === "Reception")?.id ? x.sanctioningAuthority : x.strNoteTo,
                      subject: x.noteFor === enumsObj.noteFor.find((y) => y.dValue === "Reception")?.id ? x.customerName : x.subject ? x.subject.replace(/"/g, '""') : "",
                      departmentName: x.noteFor === enumsObj.noteFor.find((y) => y.dValue === "Reception")?.id ? x.strProcessingDepartment : x.departmentName,
                      strNoteStatus: x.strNoteStatus?.trim() ? x.strNoteStatus : x.strReceptionStatus || "",
                      strNatureOfRequest: x.noteFor === enumsObj.noteFor.find((y) => y.dValue === "Reception")?.id ? x.strNatureOfRequest : x.strNatureofNote,
                      modifiedDate: x.modifiedDate ? new DateObject(new Date(x.modifiedDate)).format("DD-MMM-YYYY hh:mm:ss A") : "",
                      createdDate: x.createdDate ? new DateObject(new Date(x.createdDate)).format("DD-MMM-YYYY hh:mm:ss A") : ""
                    }))}
                    headers={exportCSVHeader()}
                  >
                    Export CSV
                  </CSVLink>
                </Button>
              </div>
            </GridToolbar>
            {renderColumnsWithData(dataResult)}
          </Grid>
        )}
      </div>
      {showNotification && (
        <Dialog
          title={<CustomDialogTitleBar />}
          onClose={() => setShowNotification(false)}
        >
          <p className="dialogcontent_">{showNotificationMsg}</p>
          <DialogActionsBar>
            <Button
              className="notifyDailogOkBtn k-button k-button-md k-rounded-md k-button-solid k-button-solid-base"
              onClick={() => setShowNotification(false)}
            >
              <span className="k-icon k-font-icon  k-i-redo cursor allIconsforPrimary-btn"></span>
              Ok
            </Button>
          </DialogActionsBar>
        </Dialog>
      )}
      <div className="pgFooterContainer">
        <Footer />
      </div>
    </div>
  );
};