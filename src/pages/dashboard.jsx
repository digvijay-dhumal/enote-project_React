import React, { useEffect, useState } from "react";
// import kendo components
import { process } from "@progress/kendo-data-query";
import { Input } from "@progress/kendo-react-inputs";
import { Button } from "@progress/kendo-react-buttons";
import {
  Grid,
  GridColumn as Column,
  GridToolbar,
} from "@progress/kendo-react-grid";
import { setGroupIds } from "@progress/kendo-react-data-tools";
// import external components
import { CSVLink } from "react-csv";
import { Link } from "react-router-dom";
import { useMsal } from "@azure/msal-react";
import DateObject from "react-date-object";
// import internal components
import { ColumnMenu } from "./custom-cells";
import { fetchNoteData } from "./apiService";
// import css styles
import "../styles/passcode.css";
import { IB_eReceptionDomain_URL } from "../config";
import { DatePicker } from "@progress/kendo-react-dateinputs";
import { MultiSelect } from "@progress/kendo-react-dropdowns"; // Changed from DropDownList to MultiSelect

const DATA_ITEM_KEY = "id";
const SELECTED_FIELD = "selected";
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

// Helper function to check if a date is in current month
const isDateInCurrentMonth = (date) => {
  if (!date) return false;
  const now = new Date();
  const inputDate = new Date(date);
  return (
    inputDate.getMonth() === now.getMonth() &&
    inputDate.getFullYear() === now.getFullYear()
  );
};


export const Dashboard = ({ apiOutput, selectedView, enumObj }) => {
  const { accounts } = useMsal();
  const [filterValue, setFilterValue] = useState("");
  const [filteredData, setFilteredData] = useState(apiOutput);
  const [dataState, setDataState] = useState(initialDataState);
  const [dataResult, setDataResult] = useState(
    process(filteredData, dataState)
  );
  const [apiData, setApiData] = useState([]);
  // Changed to arrays for multiple selections
  const [selectedMonths, setSelectedMonths] = useState([]);
  const [selectedYears, setSelectedYears] = useState([]);

  const months = [
    { text: 'January', value: 0 },
    { text: 'February', value: 1 },
    { text: 'March', value: 2 },
    { text: 'April', value: 3 },
    { text: 'May', value: 4 },
    { text: 'June', value: 5 },
    { text: 'July', value: 6 },
    { text: 'August', value: 7 },
    { text: 'September', value: 8 },
    { text: 'October', value: 9 },
    { text: 'November', value: 10 },
    { text: 'December', value: 11 }
  ];

  // Generate years from 2020 to current year + 5
  const generateYears = () => {
    const currentYear = new Date().getFullYear();
    const years = [];
    for (let year = 2020; year <= currentYear + 5; year++) {
      years.push({ text: year.toString(), value: year });
    }
    return years;
  };

  const yearsList = generateYears();

  // Set default current month and year when component mounts or view changes
  useEffect(() => {
    if (selectedView === "My Approved Notes") {
      const currentDate = new Date();
      const currentMonth = currentDate.getMonth();
      const currentYear = currentDate.getFullYear();
      
      // Set default month (as array for MultiSelect)
      const defaultMonth = months.find(month => month.value === currentMonth);
      setSelectedMonths(defaultMonth ? [defaultMonth] : []);
      
      // Set default year (as array for MultiSelect)
      const defaultYear = yearsList.find(year => year.value === currentYear);
      setSelectedYears(defaultYear ? [defaultYear] : []);
    } else {
      // Clear filters for other views
      clearFilters();
    }
  }, [selectedView]);

  useEffect(() => {
    setApiData(apiOutput);
    setFilteredData(apiOutput);

    let processedData = process(apiOutput, initialDataState);
    setDataResult(processedData);
    setDataState({ ...dataState, total: apiOutput.length });
    setFilterValue("");
  }, [apiOutput]);

  useEffect(() => {
    const normalizedData = apiOutput.map((item) => ({
      ...item,
      strNoteStatus: item.strNoteStatus?.trim()
        ? item.strNoteStatus
        : item.strReceptionStatus || "",
      departmentName: item.departmentName?.trim()
        ? item.departmentName
        : item.strProcessingDepartment || "",
      subject: item.subject?.trim()
        ? item.subject
        : item.customerName || "",
      strNatureOfRequest: item.strNatureOfRequest?.trim()
        ? item.strNatureOfRequest
        : item.strNatureofNote || "",
      strNoteTo: item.strNoteTo?.trim()
        ? item.strNoteTo
        : item.sanctioningAuthority || "",
    }));

    setApiData(normalizedData);
    
    // Apply filters ONLY if it's "My Approved Notes" view
    let dataToFilter = normalizedData;
    if (selectedView === "My Approved Notes") {
      dataToFilter = applyMonthYearFilter(normalizedData, selectedMonths, selectedYears);
    } else {
      // For other views, use the original normalized data without month/year filtering
      dataToFilter = normalizedData;
    }

    setFilteredData(dataToFilter);

    let processedData = process(dataToFilter, initialDataState);
    setDataResult(processedData);
    setDataState({ ...dataState, total: dataToFilter.length });
    setFilterValue("");
  }, [apiOutput, selectedMonths, selectedYears, selectedView]);

  // Updated applyMonthYearFilter to handle arrays
  const applyMonthYearFilter = (data, monthsArray, yearsArray) => {
    // If not "My Approved Notes", return original data
    if (selectedView !== "My Approved Notes") {
      return data;
    }

    const selectedMonthValues = monthsArray.map(m => m.value);
    const selectedYearValues = yearsArray.map(y => y.value);

    // If no filters selected, show current month and year data by default
    if (selectedMonthValues.length === 0 && selectedYearValues.length === 0) {
      const currentDate = new Date();
      const currentMonth = currentDate.getMonth();
      const currentYear = currentDate.getFullYear();
      
      return data.filter(item => {
        // Check both modifiedDate and createdDate
        const modifiedDate = item.modifiedDate ? new Date(item.modifiedDate) : null;
        const createdDate = item.createdDate ? new Date(item.createdDate) : null;

        // If neither date exists, exclude the item
        if (!modifiedDate && !createdDate) return false;

        const modifiedMonth = modifiedDate ? modifiedDate.getMonth() : null;
        const modifiedYear = modifiedDate ? modifiedDate.getFullYear() : null;
        const createdMonth = createdDate ? createdDate.getMonth() : null;
        const createdYear = createdDate ? createdDate.getFullYear() : null;

        // Check if EITHER date matches current month and year
        return (
          (modifiedMonth === currentMonth && modifiedYear === currentYear) ||
          (createdMonth === currentMonth && createdYear === currentYear)
        );
      });
    }

    // Filter data based on selected months and years
    return data.filter(item => {
      const modifiedDate = item.modifiedDate ? new Date(item.modifiedDate) : null;
      const createdDate = item.createdDate ? new Date(item.createdDate) : null;

      if (!modifiedDate && !createdDate) return false;

      const modifiedMonth = modifiedDate ? modifiedDate.getMonth() : null;
      const modifiedYear = modifiedDate ? modifiedDate.getFullYear() : null;
      const createdMonth = createdDate ? createdDate.getMonth() : null;
      const createdYear = createdDate ? createdDate.getFullYear() : null;

      // Check if item matches ANY of the selected months AND ANY of the selected years
      let matchesMonths = selectedMonthValues.length === 0;
      let matchesYears = selectedYearValues.length === 0;

      // Check month matches
      if (selectedMonthValues.length > 0) {
        matchesMonths = (
          (modifiedMonth !== null && selectedMonthValues.includes(modifiedMonth)) ||
          (createdMonth !== null && selectedMonthValues.includes(createdMonth))
        );
      }

      // Check year matches
      if (selectedYearValues.length > 0) {
        matchesYears = (
          (modifiedYear !== null && selectedYearValues.includes(modifiedYear)) ||
          (createdYear !== null && selectedYearValues.includes(createdYear))
        );
      }

      // Return true if matches both conditions (or if one condition is not specified)
      return matchesMonths && matchesYears;
    });
  };

  // Month picker change handler for MultiSelect
  const onMonthChange = (event) => {
    setSelectedMonths(event.value);
  };

  // Year picker change handler for MultiSelect
  const onYearChange = (event) => {
    setSelectedYears(event.value);
  };

  // Clear all filters
  const clearFilters = () => {
    setSelectedMonths([]);
    setSelectedYears([]);
    setFilterValue("");
    
    // If it's "My Approved Notes" view, set back to current month/year
    if (selectedView === "My Approved Notes") {
      const currentDate = new Date();
      const currentMonth = currentDate.getMonth();
      const currentYear = currentDate.getFullYear();
      
      const defaultMonth = months.find(month => month.value === currentMonth);
      setSelectedMonths(defaultMonth ? [defaultMonth] : []);
      
      const defaultYear = yearsList.find(year => year.value === currentYear);
      setSelectedYears(defaultYear ? [defaultYear] : []);
    }
  };

  // Filter data based on search text for all columns 
  const onFilterChange = (ev) => {
    let value = ev.value;
    setFilterValue(ev.value);

    // Start with base data (already filtered by month/year if applicable)
    let baseData = apiData;
    
    // For "My Approved Notes", apply month/year filters first
    if (selectedView === "My Approved Notes") {
      baseData = applyMonthYearFilter(apiData, selectedMonths, selectedYears);
    }

    if (!value) {
      // No search value - use the base data (which may be filtered by month/year)
      setFilteredData(baseData);
      setDataResult(
        process(baseData, {
          ...dataState,
          total: baseData.length,
        })
      );
    } else {
      // Apply search filter on top of base data
      let newData = baseData.filter((item) => {
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
      let clearedPagerDataState = { ...dataState, take: 10, skip: 0 };
      let processedData = process(newData, clearedPagerDataState);
      setDataResult(processedData);
      setDataState({ ...dataState, total: newData.length });
    }
  };

  const [resultState] = React.useState(
    processWithGroups(
      apiOutput.map((item) => ({
        ...item,
      })),
      initialDataState
    )
  );

  // Grid dataStateChange handler
  const dataStateChange = (event) => {
    setDataResult(process(filteredData, event.dataState));
    setDataState(event.dataState);
  };

  // Grid Expand change handler
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

  // CSV column headers
  const exportCSVHeader = (selectedView) => {
    const headers = [
      { key: "noteNumber", label: "Note ID" },
      { key: "createdByName", label: "Requester" },
      { key: "departmentName", label: "Department" },
      { key: "strNoteTo", label: "Note To", width: 200 },
      { key: "strNatureOfRequest", label: "Nature of Request", width: 200 },
      { key: "subject", label: "Subject" },
      { key: "strNoteStatus", label: "Status" },
      ...(selectedView === "My Recommended/Referred Notes" ||
        selectedView === "ED/MD Notes"
        ? [{ key: "currentActionerName", label: "Current Approver" }]
        : []),
      { key: "lastActioner", label: "Previous Approver" },
      { key: "finalApprover", label: "Final Approver" },
      { key: "modifiedDate", label: "Modified Date" },
      { key: "createdDate", label: "Created Date" },
    ];

    return headers;
  };

  // Render month and year pickers for "My Approved Notes" - updated for MultiSelect
  const renderDateFilters = () => {
    if (selectedView !== "My Approved Notes") return null;

    return (
      <div className="date-filters-container" style={{ 
        borderRadius: '4px',
        display: 'flex',
        gap: '15px',
        alignItems: 'center'
      }}>
        <div className="filter-group" style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
          <label style={{ marginBottom: '0', fontWeight: 'bold', whiteSpace: 'nowrap' }}>
            Month:
          </label>
          <MultiSelect
            data={months}
            textField="text"
            dataItemKey="value"
            value={selectedMonths}
            onChange={onMonthChange}
            placeholder="Select Month(s)"
            style={{ width: '250px' }}
          />
        </div>
        
        <div className="filter-group" style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
          <label style={{ marginBottom: '0', fontWeight: 'bold', whiteSpace: 'nowrap' }}>
            Year:
          </label>
          <MultiSelect
            data={yearsList}
            textField="text"
            dataItemKey="value"
            value={selectedYears}
            onChange={onYearChange}
            placeholder="Select Year(s)"
            style={{ width: '200px' }}
          />
        </div>

        {(selectedMonths.length > 0 || selectedYears.length > 0) && (
          <Button
            onClick={clearFilters}
            style={{ borderColor: " rgba(0, 0, 0, 0.08)", backgroundColor: "#f5f5f5", color: "#424242" }}
            themeColor="primary"
          >
            Clear Filters
          </Button>
        )}
      </div>
    );
  };

  // Based on Data column with will render
  const renderColumnsWithData = (data) => {
    if (!data || data.length === 0) {
      return null;
    }

    const getColumnsConfig = () => {
const baseColumns = [
  { field: "noteNumber", title: "Note ID", width: 120 },
  { field: "createdByName", title: "Requester", width: 180 },
  { field: "departmentName", title: "Department", width: 220 },
  { field: "strNoteTo", title: "Note To", width: 180 },
  { field: "strNatureOfRequest", title: "Nature of Request", width: 220 },
  { field: "subject", title: "Subject / Customer Name", width: 300 }, 
  { field: "strNoteStatus", title: "Status", width: 150 },
  { field: "lastActioner", title: "Previous Approver", width: 200 },
  { field: "finalApprover", title: "Final Approver", width: 200 },
  { field: "modifiedDate", title: "Modified Date", width: 180 },
  { field: "createdDate", title: "Created Date", width: 180 },
];


      // Add currentActionerName only if the selected view is "My Recommended/Referred Notes"
      if (
        selectedView === "My Recommended/Referred Notes" ||
        selectedView === "ED/MD Notes"
      ) {
        baseColumns.splice(6, 0, {
          field: "currentActionerName",
          title: "Current Approver",
          width: 150,
        });
      }

      return baseColumns;
    };
    const columnsConfig = getColumnsConfig();

    return columnsConfig.map((column) => (
      <Column
        key={column.field}
        field={column.field}
        title={column.title}
        width={column.width || undefined}
        cell={(props) => {
          const { dataItem } = props;

          const { noteFor, status, createdBy, noteId, ereceptionId } = dataItem;
          // Defensive check for enumObj and NoteStatus
          let isDraftOrReturned = false;

          if (enumObj && Array.isArray(enumObj.NoteStatus)) {
            const statusObj = enumObj.NoteStatus.find((x) => x.id === status);
            console.log(enumObj?.NoteStatus, "enumObj.NoteStatus");

            if (statusObj) {
              if (statusObj.dValue === "Returned") {
                isDraftOrReturned = true; // always editable
              } else if (
                ["Draft", "Called Back"].includes(statusObj.dValue) &&
                createdBy === accounts[0].username
              ) {
                isDraftOrReturned = true;
              }
            }
          }
          console.log("isDraftOrReturned", isDraftOrReturned);

          // Determine link path
          let linkPath = "#";
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

            default:
              linkPath = "#";
          }

          // Link styling based on field
          const linkProps =
            column.field === "noteNumber"
              ? { className: "_note_number", style: { color: "red" }, target: "_blank" }
              : { style: { color: "#424242" } };

          // Resolve raw value
          let rawValue;

          if (column.title.includes("Date")) {
            try {
              const date = new Date(dataItem[column.field]);
              rawValue = new DateObject(date).format("DD-MMM-YYYY hh:mm A");
            } catch {
              rawValue = dataItem[column.field];
            }
          }
          else if (column.field === "strNoteStatus") {
            rawValue = dataItem["strNoteStatus"]?.trim()
              ? dataItem["strNoteStatus"]
              : dataItem["strReceptionStatus"] || "";
          } else if (column.field === "departmentName") {
            rawValue = dataItem["departmentName"]?.trim()
              ? dataItem["departmentName"]
              : dataItem["strProcessingDepartment"] || "";
          } else if (column.field === "subject") {
            rawValue = dataItem["subject"]?.trim()
              ? dataItem["subject"]
              : dataItem["customerName"] || "";
          } else if (column.field === "strNatureOfRequest") {
            rawValue = dataItem["strNatureOfRequest"]?.trim()
              ? dataItem["strNatureOfRequest"]
              : dataItem["strNatureofNote"] || "";
          } else if (column.field === "strNoteTo") {
            rawValue = dataItem["strNoteTo"]?.trim()
              ? dataItem["strNoteTo"]
              : dataItem["sanctioningAuthority"] || "";
          }
          else {
            rawValue = dataItem[column.field];
          }

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

  //  Grid Header cell handler
  const onHeaderCellClick = (event) => {
    const clickedColumn = event.column;
    if (clickedColumn.field === "noteId") {
      // Extract noteId from the clicked column
      const noteId = clickedColumn.dataItem.noteId;
      // Fetch data based on the noteId
      fetchNoteData(noteId);
    }
  };

  return (
    <div>
      <Grid
        className="cstGridlandPgStyles"
        onHeaderCellClick={onHeaderCellClick}
        pageable={{ pageSizes: true }}
        data={dataResult}
        sortable={true}
        total={resultState.total}
        onDataStateChange={dataStateChange}
        {...dataState}
        onExpandChange={onExpandChange}
        expandField="expanded"
        dataItemKey={DATA_ITEM_KEY}
        selectedField={SELECTED_FIELD}
        size={"small"}
        resizable={true}
      >
        <GridToolbar>
          <div className="searchValidationCss" style={{ 
            display: 'flex', 
            justifyContent: 'space-between', 
            alignItems: 'center', 
            width: '100%' 
          }}>
            {/* Left side - Search and Export */}
            <div style={{ display: 'flex', gap: '15px', alignItems: 'center' }}>
              <div className="searchValidationCss">
                <Input
                  value={filterValue}
                  onChange={onFilterChange}
                  className="searchCSS"
                  placeholder="Search in all columns..."
                />
              </div>

              <div className="export-btns-container">
                <Button className="_exportdiv">
                  <CSVLink
                    filename={`eNote-${selectedView.replace(
                      / /g,
                      ""
                    )}${new DateObject(new Date()).format("DD-MMM-YYYY hh:mm A")}`}
                    data={filteredData.map((x) => ({
                      ...x,
                      subject:
                        x.subject?.trim()
                          ? x.subject
                          : x.customerName || "",
                      strNoteStatus:
                        x.strNoteStatus?.trim()
                          ? x.strNoteStatus
                          : x.strReceptionStatus || "",
                      departmentName:
                        x.departmentName?.trim()
                          ? x.departmentName
                          : x.strProcessingDepartment || "",
                      strNatureOfRequest:
                        x.strNatureOfRequest?.trim()
                          ? x.strNatureOfRequest
                          : x.strNatureofNote || "",
                      strNoteTo:
                        x.strNoteTo?.trim()
                          ? x.strNoteTo
                          : x.sanctioningAuthority || "",
                      modifiedDate: new DateObject(new Date(x.modifiedDate)).format("DD-MMM-YYYY hh:mm A"),
                      createdDate: new DateObject(new Date(x.createdDate)).format("DD-MMM-YYYY hh:mm A"),
                    }))}
                    headers={exportCSVHeader(selectedView)}
                  >
                    Export CSV
                  </CSVLink>
                </Button>
              </div>
            </div>

            {/* Right side - Date Filters (only for My Approved Notes) */}
            {renderDateFilters()}
          </div>
        </GridToolbar>
        {renderColumnsWithData(dataResult)}
      </Grid>
    </div>
  );
};