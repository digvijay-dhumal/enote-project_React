import React, { useState, useEffect } from "react";
import { Grid, GridColumn as Column, GridToolbar } from "@progress/kendo-react-grid";
import { orderBy } from "@progress/kendo-data-query";
import { process } from "@progress/kendo-data-query";
import { Input } from "@progress/kendo-react-inputs";
import { setGroupIds } from "@progress/kendo-react-data-tools";
import { getter } from "@progress/kendo-react-common";
// import external components
import { useMsal, useAccount } from "@azure/msal-react";
import DateObject from "react-date-object";
// import internal components
import { API_BASE_URL, API_ENDPOINTS, loginRequest, API_COMMON_HEADERS,charValidation,sanitizeInput  } from "../config";
import { PageLoader } from "./pageLoader";
import { getAccessToken } from "../App";
import { ColumnMenu } from "../pages/custom-cells";
// import css styles
import "../styles/datagridpage.css";
import "../styles/forms.css";
const initialSort = [
  {
    field: "modifiedDate",//noteId
    dir: "desc",
  },
];
const DATA_ITEM_KEY = "noteId";
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

const AdminDashboardGrid = (props) => {
  const idGetter = getter(DATA_ITEM_KEY);
  const [filterValue, setFilterValue] = React.useState("");
  const [filteredData, setFilteredData] = React.useState();
  const [currentSelectedState, setCurrentSelectedState] = React.useState({});
  const [dataState, setDataState] = React.useState(initialDataState);
  const [data, setData] = React.useState(filteredData);
  const [apiData, setApiData] = React.useState([]);
  const [dataResult, setDateGrid] = useState([]);
  const [isLoading, setIsLoading] = useState(true);
  const { accounts, instance } = useMsal();
  const account = useAccount(accounts[0] || {});
  const [filterErrorMessage, setFilterErrorMessage] = useState("");

  useEffect(() => {
    if (props.departmentName) {
      fetchApiData();
    }
  }, [props]);

  const fetchApiData = async () => {
    setFilterValue('');
    if (props.departmentName) {
      const Params = {
        "Department": (props.departmentName).replace(/&/g, "%26"),
        "fromDate": props.fromDate ? new DateObject(new Date(props.fromDate)).format("DD-MM-YYYY") : null,
        "ToDate": props.toDate ? new DateObject(new Date(props.toDate)).format("DD-MM-YYYY") : null,
      };
      try {
        const accessToken = await getAccessToken({ ...loginRequest, account }, instance);
        const response = await fetch(
          `${API_BASE_URL}${API_ENDPOINTS.GET_DeatilsForPending}?Department=${Params.Department}&fromDate=${Params.fromDate}&ToDate=${Params.ToDate}`,
          {
            method: "POST",
            headers: { ...API_COMMON_HEADERS, Authorization: `Bearer ${accessToken}` },
          }
        );

        if (response.ok) {
          const apiOutput = await response.json();
          const resData = apiOutput?.map((x, ind) => ({ ...x, SerialNo: ind + 1 }));
          const apiData = orderBy(resData, initialSort);
          if (Array.isArray(apiData)) {
            setApiData(apiData);
            setFilteredData(apiData);
            setData(apiData);
            setDateGrid(process(apiData, dataState));
            setDataState({ ...dataState, total: apiData.length });
          } else {
            console.error("Error fetching data: Invalid API response", apiData);
            setFilteredData([]);
            setDateGrid(process([], dataState));
          }
        } else {
          console.error("Error fetching data:", response.statusText);
        }
      } catch (error) {
        console.error("Error fetching data:", error);
      } finally {
        setIsLoading(false);
      }
    }
  };

  const onFilterChange = (ev) => {
    let value = ev.value;
  
    // Sanitize input to remove invisible characters
    // value = sanitizeInput(value);
  
    // // Allowed characters regex
    // const allowedCharsRegex = charValidation;
  
    // // Replace invalid characters with an empty string
    // if (!allowedCharsRegex.test(value)) {
    //   setFilterErrorMessage('Special characters are not allowed.');
    //   value = value.replace(/[^a-zA-Z0-9.,:;?!_@#₹ \t\n&"'\-\$%]/g, '');
    //   setFilteredData(apiData);
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
      setDateGrid(newDataResult);
  
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
      setDateGrid(processedData);
      setDataState({ ...clearedPagerDataState, total: newData.length });
      setData(newData);
    }
  };
  

  const [resultState, setResultState] = useState(
    processWithGroups(
      apiData.map((item) => ({
        ...item,
        ["selected"]: currentSelectedState[idGetter(item)],
      })),
      initialDataState
    )
  );
  const dataStateChange = (event) => {
    setDateGrid(process(filteredData, event.dataState));
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
          setDateGrid({
            ...dataResult,
            data: newData,
          });
        }
      } else {
        item.expanded = event.value;
        setDateGrid({
          ...dataResult,
          data: newData,
        });
      }
    },
    [dataResult]
  );
 
  const renderColumnsWithData = (data) => {
    if (!data || data.length === 0) {
      return null;
    }

    let columnsConfig = [];

    columnsConfig = [
      { field: "SerialNo", title: "Sl.No", minWidth: 15, },
      { field: "currentActionerName", title: "Pending with", minWidth: 30, },
      { field: "designation", title: "Designation ", minWidth: 25 },
      { field: "srNo", title: "SrNo", minWidth: 20, },
      { field: "count", title: "Count", minWidth: 10 },
    ];

    return columnsConfig.map((column) => (
      <Column
        key={column.field}
        field={column.field}
        title={column.title}
        columnMenu={ColumnMenu}
      />
    ));
  };

  return (
    <div className="largeScreen">
      {isLoading ? (
     <div>
          {/* <PageLoader /> */}
          </div>
      ) : (
        <Grid
          className="cstGridStylesForAdmin"
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
                className='searchCSS'
                placeholder="Search in all columns..."
              />
              {filterErrorMessage && (
                <div className="inCorrectFileError">
                  {filterErrorMessage}
                </div>
              )}
            </div>

            <div className="export-btns-container">
            </div>
          </GridToolbar>
          {renderColumnsWithData(dataResult)}
        </Grid>
      )}
    </div>
  );
};

export default AdminDashboardGrid;