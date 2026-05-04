import React, { useState, useEffect } from "react";
import {
  Grid,
  GridColumn as Column,
  GridToolbar,
} from "@progress/kendo-react-grid";
import { Link } from "react-router-dom";
import { Button } from "@progress/kendo-react-buttons";
import { DropDownList } from "@progress/kendo-react-dropdowns";
import { Dialog, DialogActionsBar } from "@progress/kendo-react-dialogs";
import Navbar from "../components/navbar";
import Footer from "../components/footer";
import { useParams } from "react-router";
import { process } from "@progress/kendo-data-query";
import { Input } from "@progress/kendo-react-inputs";
import { CSVLink } from "react-csv";
import { useMsal, useAccount } from "@azure/msal-react";
import { ColumnMenu } from "./custom-cells";
import { setGroupIds, setExpandedState } from "@progress/kendo-react-data-tools";
import { Sidebar } from "../components/sidebar";
import { MenuNavContainer } from "../components/menu";
import { API_BASE_URL, API_ENDPOINTS, charValidation, sanitizeInput } from "../config";
import { PageLoader } from "../components/pageLoader";
import "../styles/datagridpage.css";
import "../styles/forms.css";
import DateObject from "react-date-object";
import { orderBy } from "@progress/kendo-data-query";
import { getAccessToken } from "../App";
import { loginRequest } from "../config";
import { API_COMMON_HEADERS } from "../config";

const initialSort = [
  {
    field: "modifiedDate",
    dir: "desc",
  },
];

const DATA_ITEM_KEY = "id";
const SELECTED_FIELD = "Select";
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

const ATRViews = () => {
  const idGetter = (item) => item.id;
  const { id } = useParams();
  const [filterValue, setFilterValue] = React.useState("");
  const [filteredData, setFilteredData] = React.useState([]);
  const [currentSelectedState, setCurrentSelectedState] = React.useState({});
  const [dataState, setDataState] = React.useState(initialDataState);
  const [data, setData] = React.useState(filteredData);
  const [apiData, setApiData] = React.useState([]);
  const [dataResult, setDataResult] = useState([]);
  const [isLoading, setIsLoading] = useState(true);
  const [currentHeading, setCurrentHeading] = useState();
  const { accounts, instance } = useMsal();
  const [selectedView, setSelectedView] = useState("");
  const account = useAccount(accounts[0] || {});
  const [filterErrorMessage, setFilterErrorMessage] = useState("");
  const [ATRClosureDialog, setATRClosureDialog] = useState(false);
  const [showErrorMessage, setShowErrorMessage] = useState("");
  const [ATRClosureDeleteDialog, setATRClosureDeleteDialog] = useState(false);
  const [enymObj, setenymObj] = useState({});


  const options = [
    { text: "Not an action point", key: "Not an action point" },
  ];
  const [selectedDropdownValue, setDropdownSelectedValue] = useState(options[0]);

  const handleDropdownChange = (event) => {
    setDropdownSelectedValue(event.value);
  };

  const switchNoteId = (id) => {
    setCurrentHeading(id);
    setIsLoading(true);
    switch (id) {

      case "Pending ATR":
        fetchApiData(API_ENDPOINTS.ATR_GetRequestsByStatus, "Submitted", "pendingAtr");
        setSelectedView(id);
        break;
      case "Pending Approval":
        fetchApiData(API_ENDPOINTS.ATR_GetRequestsByStatus, "Action Taken", "pendingAtrSectt");
        setSelectedView(id);
        break;
      case "Completed ATR":
        fetchApiData(API_ENDPOINTS.ATR_GetRequestsByStatus, "Completed", "completedAtr");
        setSelectedView(id);
        break;
      case "ATR Closure":
        fetchApiData(API_ENDPOINTS.ATR_GetRequestsByStatus, "Completed", "atrClosure");
        setSelectedView(id);
        break;
      case "All ATR":
        fetchApiDataAll();
        setSelectedView(id);
        break;
      default:
        fetchApiData(API_ENDPOINTS.ATR_GetRequestsByStatus, "Submitted", "pendingAtr");
        setSelectedView("Pending ATR");
        return 1;
    }
  };

  useEffect(() => {
    switchNoteId(id); // Set the default noteId
  }, [id]);

  const fetchApiData = async (endPoint, statusTxt, type) => {
    setFilterValue('');

    const accessToken = await getAccessToken({ ...loginRequest, account }, instance);

    const dropdowns = await fetch(`${API_BASE_URL}${API_ENDPOINTS.GET_DROPDOWNDATA}`, {
      method: "GET",
      headers: { ...API_COMMON_HEADERS, Authorization: `Bearer ${accessToken}` }
    });

    const enumsObj = await dropdowns.json();
    setenymObj(enumsObj);
    const endpoint = `${API_BASE_URL}${endPoint}`;

    // Set AtrStatus based on condition (ATR Closure should always be 1)
    const atrStatus = (statusTxt === "Completed" && type === "atrClosure")
      ? 1
      : enumsObj.ATRStatusEnum.find((x) => x.dValue === statusTxt)?.id;

    if (!atrStatus) {
      console.error(`Error: Status '${statusTxt}' not found in ATRStatusEnum.`);
      return;  // Early exit if status is not found
    }

    try {
      const response = await fetch(endpoint, {
        method: "POST",
        body: JSON.stringify({
          AtrStatus: atrStatus, // Use the resolved AtrStatus
          LoginMailId: accounts[0].username,
          ATRReportType: type,
        }),
        headers: { ...API_COMMON_HEADERS, Authorization: `Bearer ${accessToken}` },
      });

      if (response.ok) {
        const apiObj = await response.json();
        const resData = apiObj.pendingATRsList;
        const apiCstData = orderBy(resData, initialSort);

        if (apiObj.statusMessage === "Success") {
          setApiData(apiCstData);
          setFilteredData(apiCstData);
          setData(apiCstData);
          setDataResult(process(apiCstData, dataState));
          setDataState({ ...dataState, total: apiCstData.length });
        } else {
          console.error("Error fetching data: Invalid API response", apiCstData);
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

  const fetchApiDataAll = async () => {
    setFilterValue('');

    const accessToken = await getAccessToken({ ...loginRequest, account }, instance);

    try {
      const response = await fetch(
        `${API_BASE_URL}${API_ENDPOINTS.ATR_GetAllRequests}`,
        {
          method: "Post",
          body: JSON.stringify({
            LoginMailId: accounts[0].username,
          }),
          headers: { ...API_COMMON_HEADERS, Authorization: `Bearer ${accessToken}` }
        }
      );

      if (response.ok) {
        const apiObj = await response.json();

        const resData = apiObj.pendingATRsList;
        const apiCstData = orderBy(resData, initialSort);

        if (apiObj.statusMessage === "Success") {
          setApiData(apiCstData);
          setFilteredData(apiCstData);
          setData(apiCstData);
          setDataResult(process(apiCstData, dataState));
          setDataState({ ...dataState, total: apiCstData.length });
        } else {
          console.error(
            "Error fetching data: Invalid API response",
            apiCstData
          );
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

    // Sanitize input to remove invisible characters
    // value = sanitizeInput(value);

    // // Allowed characters regex
    // const allowedCharsRegex = charValidation;

    // Validate characters after sanitization
    // if (!allowedCharsRegex.test(value)) {
    //   setFilterErrorMessage('Special characters are not allowed.');
    //   value = value.replace(/[^a-zA-Z0-9.,:;?!_@#₹ \t\n&"'\-\$%]/g, '');
    //   setFilteredData(apiData); // Reset to original if invalid
    // } else {
    //   setFilterErrorMessage('');
    // }

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
            item[property].toString().toLowerCase().includes(value.toLowerCase())
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

      const newDataResult = processWithGroups(newData, dataState);
      setDataResult(newDataResult);

      setDataState((prevDataState) => ({
        ...prevDataState,
        total: newData.length,
      }));

      let clearedPagerDataState = {
        ...dataState,
        take: 10,
        skip: 0,
      };

      let processedData = process(newData, clearedPagerDataState);
      setDataResult(processedData);
      setDataState({ ...clearedPagerDataState, total: newData.length });
      setData(newData);
    }
  };


  const [resultState, setResultState] = React.useState(
    processWithGroups(
      apiData.map((item) => ({
        ...item,
        ["selected"]: currentSelectedState[idGetter(item)],
      })),
      initialDataState
    )
  );
  const dataStateChange = (event) => {
    if (event.dataState && filteredData) {
      setDataResult(process(filteredData, event.dataState));
      setDataState(event.dataState);
    }
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
          ["selected"]: currentSelectedState[idGetter(item)],
        };
      }
    });
    return newData;
  };

  const newData = setExpandedState({
    data: setSelectedValue(resultState.data),
    collapsedIds: [],
  });

  const getNumberOfItems = (data) => {
    let count = 0;
    data.forEach((item) => {
      if (item.items) {
        count = count + getNumberOfItems(item.items);
      } else {
        count++;
      }
    });
    return count;
  };
  const getNumberOfSelectedItems = (data) => {
    let count = 0;
    data.forEach((item) => {
      if (item.items) {
        count = count + getNumberOfSelectedItems(item.items);
      } else {
        count = count + (item.selected === true ? 1 : 0);
      }
    });
    return count;
  };

  const exportCSVHeader = () => {
    return [
      { key: "noteNumber", label: "Note#" },
      { key: "atrAssignerEmailName", label: "Assignee" },
      { key: "departmentName", label: "Department" },
      { key: "approverEmailName", label: "Assigned By" },
      { key: "strAtrStatus", label: "Status" },
      { key: "subject", label: "Subject" },
      { key: "remarks", label: "Remarks" },
      { key: "createdDate", label: "Created Date" },
    ];
  }
  const checkHeaderSelectionValue = () => {
    // Check if all rows are selected (when selecting the header checkbox)
    let selectedItems = getNumberOfSelectedItems(newData); // Assuming getNumberOfSelectedItems counts the 'selected' rows

    // Return true if all items are selected, false otherwise
    return newData.length > 0 && selectedItems === getNumberOfItems(newData); // getNumberOfItems gets the total count of rows
  };

  // Row checkbox change handler
  const handleCheckboxChange = (event, dataItem) => {
    const newData = data.map((item) => {
      if (item.atrId === dataItem.atrId) {

        // Toggle the selected status of the clicked row
        return { ...item, selected: !item.selected }; // Change selected state
      }
      return item; // Return the unchanged item for others
    });

    // Update the state with the new selection status
    setData(newData);
    setDataResult(process(newData, dataState)); // Re-render the Grid with updated data
  };

  // SAVE button click handler for ATR Closure
  const handleSaveClick = async () => {
    setATRClosureDeleteDialog(false);

    // Check if a dropdown value is selected
    if (!selectedDropdownValue) {
      setATRClosureDialog(true);
      setShowErrorMessage("Please select a value from the dropdown.");
      return; // Exit if dropdown value is not selected
    }

    // Collect all selected rows' atrPrimeIDs
    const selectedItems = data.filter((item) => item.selected);
    const selectedAtrIds = selectedItems.map((item) => item.atrId);

    if (selectedItems.length === 0) {
      setATRClosureDialog(true);
      setShowErrorMessage("Please select at least one item to proceed.");
      return; // Exit if dropdown value is not selected
    }

    // Find the status ID using enumsObj.ATRStatusEnum
    const statusTxt = "Not Applicable"; // Replace this with your actual status text
    const atrStatusId = enymObj.ATRStatusEnum.find((x) => x.dValue === statusTxt)?.id || 0;

    const accessToken = await getAccessToken({ ...loginRequest, account }, instance);
    const params = {
      atrPrimeID: selectedAtrIds, // Send selected IDs as an array
      atrStatusVal: atrStatusId,  // Pass the found status ID
      actionBy: accounts[0].username,
    };

    try {
      const response = await fetch(`${API_BASE_URL}${API_ENDPOINTS.eNote_ATRBulkUpdate}`, {
        method: "POST",
        body: JSON.stringify(params),
        headers: { ...API_COMMON_HEADERS, Authorization: `Bearer ${accessToken}` },
      });

      if (response.ok) {
        const atrBulk = await response.text();
        console.log(atrBulk, "atrBulk");
        setATRClosureDialog(true);
        setShowErrorMessage("ATR closure updated successfully.");

        // Clear the checkboxes by resetting the selected state
        const clearedData = data.map((item) => ({ ...item, selected: false }));
        setData(clearedData);
        setDataResult(process(clearedData, dataState));
        setDropdownSelectedValue(null);
        // window.location.reload(); // Reload the page to refresh the data
      } else {
        console.error("Error fetching data:", response.statusText);
      }
    } catch (error) {
      console.error("Error fetching data:", error);
    } finally {
      setIsLoading(false);
    }
  };

  const handleReload = () => {
    window.location.reload();
  }

  const handleCancel = () => {
    setATRClosureDeleteDialog(false);
    const clearedData = data.map((item) => ({ ...item, selected: false }));
    setData(clearedData);
    setDataResult(process(clearedData, dataState));
  }


  const renderColumnsWithData = (data) => {
    if (!data || data.length === 0) {
      return null;
    }

    const columnsConfig = [
      { field: "noteNumber", title: "Note#", width: 200 },
      { field: "atrAssignerEmailName", title: "Assignee", width: 200 },
      { field: "departmentName", title: "Department", width: 200 },
      { field: "approverEmailName", title: "Assigned By", width: 200 },
      { field: "strAtrStatus", title: "Status", width: 200 },
      { field: "subject", title: "Subject", width: 750 },
      { field: "remarks", title: "Remarks", width: 200 },
      { field: "createdDate", title: "Created Date", width: 200 },
    ];

    return columnsConfig.map((column) => (
      <Column
        key={column.field}
        field={column.field}
        title={column.title}
        width={column.width || undefined}
        cell={(props) => {
          const cellValue = column.title.includes("Date")
            ? new DateObject(new Date(props.dataItem[column.field])).format("DD-MMM-YYYY hh:mm A")
            : props.dataItem[column.field];

          const linkProps =
            column.field === "noteNumber"
              ? { className: "_note_number", style: { color: "red" },target: "_blank" }
              : { style: { color: "#424242" } };

          return (
            <td>
              <Link to={`/eNoteAtrworkflowform/${props.dataItem["atrId"]}`} {...linkProps}>
                {cellValue}
              </Link>
            </td>
          );
        }}
        columnMenu={ColumnMenu}
      />
    ));
  };

  return (
    <div>
      <Navbar header="IB Smart Office - eNote" />
      <Sidebar />
      <MenuNavContainer isMenuPage />
      <div className="container largeScreen datagridpage">
        <div className="SectionHeads row mobileSectionHeads">{currentHeading}s</div>

        {isLoading ? (
          <div>
          {/* <PageLoader /> */}
          </div>
        ) : (
          <div>
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
              resizable={true}
              size={"small"}
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
                    <div className="inCorrectFileError">
                      {filterErrorMessage}
                    </div>
                  )}
                </div>

                <div className="export-btns-container">
                  <Button className="_exportdiv">
                    <CSVLink
                      filename={`eNote-${selectedView.replace(/ /g, "")}${new DateObject(new Date()).format("DDMMYYYYhhmmss")}`}
                      data={filteredData.map((x) => ({
                        ...x,
                        modifiedDate: new DateObject(new Date(x.modifiedDate)).format("DD-MMM-YYYY hh:mm A"),
                        createdDate: new DateObject(new Date(x.createdDate)).format("DD-MMM-YYYY hh:mm A"),
                      }))}
                      headers={exportCSVHeader()}
                    >
                      Export CSV
                    </CSVLink>
                  </Button>
                </div>
              </GridToolbar>

              {currentHeading === "ATR Closure" && (
                <Column
                  field={SELECTED_FIELD}
                  width="100px"
                  headerSelectionValue={checkHeaderSelectionValue()}
                  headerClassName="header-checkbox"
                  cell={(props) => (
                    <td>
                      <input
                        type="checkbox"
                        checked={!!props.dataItem.selected} // Ensure it's bound to the row's data
                        onChange={(e) => handleCheckboxChange(e, props.dataItem)}
                      />
                    </td>
                  )}
                />
              )}

              {renderColumnsWithData(dataResult)}
            </Grid>
            {currentHeading === "ATR Closure" && (
              <div className="atrCloseMenu">
                <div className="flex-container atrStyle">
                  <DropDownList
                    className="atrClosuredrop"
                    data={options}
                    value={selectedDropdownValue}
                    textField="text"
                    dataItemKey="key"
                    onChange={handleDropdownChange}
                  />
                  <Button
                    className="ATRsave k-button k-button-md k-rounded-md k-button-solid k-button-solid-base formBtnColor"
                    onClick={() => setATRClosureDeleteDialog(true)}
                  >
                    Save
                  </Button>
                </div>
              </div>
            )}
          </div>
        )}
      </div>

      {ATRClosureDeleteDialog && (
        <Dialog
          title="ATR Closure"
          onClose={() => setATRClosureDeleteDialog(false)}
        >
          <p className="dialogcontent_">
            Are you sure you want to close the selected ATRs?
          </p>
          <p className="dialogcontent_">
            If so, Please click on confirm button.
          </p>
          <DialogActionsBar>
            <Button className="formBtnColor" onClick={handleSaveClick}>
              <span className="k-icon k-font-icon  k-i-checkmark-circle cursor allIconsforPrimary-btn"></span>
              <span>Confirm</span>
            </Button>
            <Button onClick={handleCancel}>
              <span className="k-icon k-font-icon  k-i-close-circle cursor allIconsforPrimary-btn"></span>
              <span>Cancel</span>
            </Button>
          </DialogActionsBar>
        </Dialog>
      )}
      {ATRClosureDialog && (
        <Dialog
          title="ATR Closure"
          onClose={() => setATRClosureDialog(false)}
        >
          <p className="dialogcontent_">
            {showErrorMessage}
          </p>
          <DialogActionsBar>
            <Button onClick={handleReload} className="notifyDailogOkBtn">
              <span className="k-icon k-font-icon  k-i-redo cursor allIconsforPrimary-btn"></span>
              <span>OK</span>
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

export default ATRViews;