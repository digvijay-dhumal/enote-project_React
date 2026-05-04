import React, { useState, useEffect } from "react";
// import kendo components
import { Button } from "@progress/kendo-react-buttons";
import { orderBy } from "@progress/kendo-data-query";
import { DropDownList } from "@progress/kendo-react-dropdowns";
// import external components
import { useMsal, useAccount } from "@azure/msal-react";
import DateObject from "react-date-object";
// import internal components
import Navbar from "../components/navbar";
import Footer from "../components/footer";
import { Dashboard } from "./dashboard";
import { Sidebar } from "../components/sidebar";
import { MenuNavContainer } from "../components/menu";
import { API_BASE_URL, API_ENDPOINTS, loginRequest, API_COMMON_HEADERS, } from "../config";
import { getAccessToken, useTabContext } from "../App";
import { ChartContainer } from "../components/enoteChartInfo";
// import css styles
import "../styles/datagridpage.css";

const initialSort = [
  {
    field: "modifiedDate",
    dir: "desc",
  },
];

export const ENoteHome = () => {
  const { instance, accounts } = useMsal();
  const { activeTab, activeButtons, setIsFromUrl, isFromUrl } = useTabContext();
  const [activeButton, setActiveButton] = useState(activeTab);
  const [apiOutput, setapiOutput] = React.useState([]);
  const [isSecretary, setIsSecretary] = useState(true);
  const [selectedView, setSelectedView] = useState("");
  const [apiOutputforChart, setApiOutputforChart] = useState([]);
  const account = useAccount(accounts[0] || {});
  const [isMobileView, setIsMobileView] = useState(window.innerWidth <= 1150);
  const [selectedOption, setSelectedOption] = useState(null);
  const [pendingNoteCount, setPendingNoteCount] = useState(0);
  const [returnNoteCount, setReturnNoteCount] = useState(0);
  const [approvedNotesCount, setApprovedNotesCount] = useState(0);
  const [getRefferedNoteCount, setGetRefferedNoteCount] = useState(0);
  const [edMdNotesCount, setEdMdNotesCount] = useState(0);
  const [upcomingNotesCount, setUpcomingNotesCount] = useState(0);
  const [enumObj, setEnumObj] = useState();

  const fetchAndSetData = async () => {
    try {
      const accessToken = await getAccessToken({ ...loginRequest, account }, instance);
      const response = await fetch(`${API_BASE_URL}${API_ENDPOINTS.eNote_GetNoteChartData}?userEmailId=${accounts[0].username}`, {
        method: "GET",
        headers: {
          ...API_COMMON_HEADERS,
          Authorization: `Bearer ${accessToken}`,
        },
      });

      if (response.ok) {
        const apiOutput = await response.json();

        // counts
        setPendingNoteCount(apiOutput.pendingNotesCount);
        setReturnNoteCount(apiOutput.returnedNotesCount);
        setApprovedNotesCount(apiOutput.approvedNotesCount);
        setGetRefferedNoteCount(apiOutput.refferedNoteCount);
        setUpcomingNotesCount(apiOutput.upcomingRecordCount);
        setEdMdNotesCount(apiOutput.edMdNotesCount);

        // call existing logic
        handleButtonClick(activeTab);
        findIsSecratary();

        const initialOption = options.find(
          (option) => option.value === activeTab
        );
        setSelectedOption(initialOption);
      }
    } catch (error) {
      console.error("Error fetching data:", error);
    }
  };

  // First load
  useEffect(() => {
    fetchAndSetData();

    const handleResize = () => {
      setIsMobileView(window.innerWidth <= 1150);
    };
    window.addEventListener("resize", handleResize);

    return () => window.removeEventListener("resize", handleResize);
  }, []);

  // 🔁 Auto-refresh every 10 minutes
  useEffect(() => {
    const interval = setInterval(() => {
      fetchAndSetData();
    }, 10 * 60 * 1000);

    return () => clearInterval(interval);
  }, []);

  // get data for charts
  const fetchDataForcharts = async (endPoint) => {
    // custom colors
    const colorCodes = [
      "#7731a3", // Deep Purple
      "rgb(254, 215, 76)", // Yellow
      "#ba3294", // Purple
      "rgb(78, 185, 167)", // Turquoise
      "#d92b2b", // Red
      "rgb(12, 77, 162)", // Blue
      "#4dc313", // Green
      "#1e6bb5", // Dark Blue
      "rgb(255, 165, 0)", // Orange
    ];

    try {
      const accessToken = await getAccessToken({ ...loginRequest, account }, instance);
      const response = await fetch(`${API_BASE_URL}${endPoint}?userEmailId=${accounts[0].username}`, {
        method: "GET",
        headers: { ...API_COMMON_HEADERS, Authorization: `Bearer ${accessToken}` },
      });

      if (response.ok) {
        let noteCount = 0;

        const apiOutput = await response.json();
        // adding color property to Response. adding custom colors dynamically
        // change 20/04 conacting status and count
        apiOutput.lstNoteChart = apiOutput.lstNoteChart?.map((obj, index) => {
          return {
            ...obj,
            color: colorCodes[index],
            status: `${obj.status}: ${obj.count}`,
          };
        });
        // adding color property to Response. adding custom colors dynamically
        // change 20/04 conacting status and count
        apiOutput.lstNatureofNoteStatusCount =
          apiOutput.lstNatureofNoteStatusCount?.map((obj, index) => {
            return {
              ...obj,
              color: colorCodes[index],
              status: `${obj.status}: ${obj.count}`,
            };
          });
        //Get the total count for note using count property.
        apiOutput.lstNoteChart?.map((obj) => {
          noteCount = noteCount + obj.count;
        });
        setApiOutputforChart(apiOutput);
        setPendingNoteCount(apiOutput.pendingNotesCount);
        setReturnNoteCount(apiOutput.returnedNotesCount);
        setApprovedNotesCount(apiOutput.approvedNotesCount);
        setGetRefferedNoteCount(apiOutput.refferedNoteCount);
        setEdMdNotesCount(apiOutput.edMdNotesCount);
        setUpcomingNotesCount(apiOutput.upcomingRecordCount);
      }
    } catch {
      console.log("err");
    }
  };

  // Based on button text API will call
  const handleButtonClick = async (buttonText) => {
    if (isFromUrl) {
      const url = new URL(window.location.href);
      url.searchParams.delete("tab");
      window.history.replaceState({}, "", url);
      setIsFromUrl(false);
    }

    setActiveButton(buttonText);

    const accessToken = await getAccessToken({ ...loginRequest, account }, instance);

    const dropdowns = await fetch(`${API_BASE_URL}${API_ENDPOINTS.GET_DROPDOWNDATA}`, {
      method: "GET",
      headers: { ...API_COMMON_HEADERS, Authorization: `Bearer ${accessToken}` },
    });
    const enumsObject = await dropdowns.json();
    setEnumObj(enumsObject);
console.log("buttonText:", enumsObject);
    switch (buttonText) {
      case "My Pending Notes":
        fetchApiData(API_ENDPOINTS.eNote_GetRequests, enumsObject?.NoteStatus.find((x) => x.dValue === "Pending").id, buttonText, accessToken);
        break;
      case "My Recommended/Referred Notes":
        fetchApiData(API_ENDPOINTS.eNote_GetRefferedNoteList, enumsObject?.ReferrerStatus.find((x) => x.dValue === "Pending").id, buttonText, accessToken);
        break;
      case "My Returned Notes":
        fetchApiData(API_ENDPOINTS.eNote_GetRequests, enumsObject?.NoteStatus.find((x) => x.dValue === "Returned").id, buttonText, accessToken);
        break;
      case "My Approved Notes":
        fetchApiData(API_ENDPOINTS.eNote_GetRequests, enumsObject?.NoteStatus.find((x) => x.dValue === "Approved").id, buttonText, accessToken);
        break;
      case "My Upcoming Notes":
        fetchApiData(API_ENDPOINTS.eNote_GetRequests, enumsObject?.NoteStatus.find((x) => x.dValue === "UpcomingNotes").id, buttonText, accessToken);
        break;
      case "My Notes":
        fetchDataForcharts(API_ENDPOINTS.eNote_GetNoteChartData, accessToken);
        break;
      case "ED/MD Notes":
        fetchApiData(API_ENDPOINTS.eNote_GetEDMDNoteList, 0, buttonText, accessToken);
        break;
      default:
        fetchApiData(API_ENDPOINTS.eNote_GetRequests, enumsObject?.NoteStatus.find((x) => x.dValue === "Pending").id, "My Pending Notes", accessToken);
        break;
    }
  };

  // checking if login user is Secratary or not.if login user is Secratary displaying ED/MD btn
  const findIsSecratary = async () => {
    const accessToken = await getAccessToken({ ...loginRequest, account, }, instance);

    await fetch(`${API_BASE_URL}${API_ENDPOINTS.eNote_findIsloginUserSecretary}`, {
      method: "POST",
      body: JSON.stringify({ SecretaryEmail: accounts[0].username }),
      headers: { ...API_COMMON_HEADERS, Authorization: `Bearer ${accessToken}` },
    })
      .then((res) => {
        return res.json();
      })
      .then((data) => {
        setIsSecretary(data);
      })
      .catch((error) => {
        console.error("Error fetching data:", error);
      });
  };

  // Based on button  click  this function will call
  const fetchApiData = async (apiEndpoint, Status, buttonText, accessToken) => {
    setSelectedView(buttonText);

    try {
      const response = await fetch(`${API_BASE_URL}${apiEndpoint}`, {
        method: "POST",
        body: JSON.stringify({
          status: Status,
          createdBy: accounts[0].username,
        }),
        headers: {
          ...API_COMMON_HEADERS,
          Authorization: `Bearer ${accessToken}`,
        },
      });

      if (response.ok) {
        const apiData = await response.json();
        if (apiData.pendingNoteList) {
          let resData = apiData.pendingNoteList;
          const ordData = orderBy(resData, initialSort);
          // converting date object for  modifieddate and createdDate
          const apiCstData = ordData.map((x) => ({
            ...x,
            modifiedDate: new DateObject(new Date(x.modifiedDate)).format("DD-MMM-YYYY hh:mm:ss A"),
            createdDate: new DateObject(new Date(x.createdDate)).format("DD-MMM-YYYY hh:mm:ss A"),
          }));
          setapiOutput(apiCstData);
          // based on button text  store  the total count of Response
          // Update count based on button text
          if (buttonText === "My Upcoming Notes") {
            setUpcomingNotesCount(apiData.pendingNoteList.length);
          }
        }
      } else {
        console.error("Error fetching data:", response.statusText);
      }
    } catch (error) {
      console.error("Error fetching data:", error);
    }
  };
  const options = [
    { text: "My Dashboard", value: "My Notes" },
    { text: "My Pending Notes", value: "My Pending Notes" },
    { text: "My Recommended/Referred Notes", value: "My Recommended/Referred Notes" },
    { text: "My Returned Notes", value: "My Returned Notes" },
    { text: "My Approved Notes", value: "My Approved Notes" },
    { text: "My Upcoming Notes", value: "My Upcoming Notes" }
  ];

  if (isSecretary) { options.push({ text: "ED/MD Notes", value: "ED/MD Notes" }); }

  const handleSelect = (event) => {
    const selectedValue = event.target.value;
    setSelectedOption(options.find((option) => option.value === selectedValue));
    setSelectedOption(selectedValue);
    handleButtonClick(selectedValue?.value);
  };

  const renderDropdown = () => (
    <DropDownList
      className="landing_Page_Dropdown"
      data={options}
      value={selectedOption}
      textField="text"
      dataItemKey="value"
      onChange={handleSelect}
    />
  );

  const renderButtons = () => (
    <div className="row landingPgTopBtnRow" style={{ display: 'flex', flexWrap: 'nowrap', gap: '10px' }}>
      <Button
        className="landingPgTopBtn"
        style={{
          flex: '1',
          minWidth: '0',
          background: activeButton === "My Notes" ? "#034ea1" : "",
          color: activeButton === "My Notes" ? "#ffff" : "",
        }}
        onClick={() => handleButtonClick("My Notes")}
      >
        My Dashboard
      </Button>

      <Button
        className="landingPgTopBtn"
        style={{
          flex: '1',
          minWidth: '0',
          backgroundColor: isFromUrl ? activeButtons === "My Pending Notes" ? "#034ea1" : "" : activeButton === "My Pending Notes" ? "#034ea1" : "",
          color: isFromUrl ? activeButtons === "My Pending Notes" ? "#fff" : "#333333" : activeButton === "My Pending Notes" ? "#fff" : "#333333",
        }}
        onClick={() => handleButtonClick("My Pending Notes")}
      >
        My Pending Notes
        <div className="landingPgTopBtncontent">{pendingNoteCount}</div>
      </Button>

      <Button
        className="landingPgTopBtn"
        style={{
          flex: '1',
          minWidth: '0',
          background: activeButton === "My Recommended/Referred Notes" ? "#034ea1" : "",
          color: activeButton === "My Recommended/Referred Notes" ? "#ffff" : "",
        }}
        onClick={() => handleButtonClick("My Recommended/Referred Notes")}
      >
        My Recommended/Referred Notes
        <div className="landingPgTopBtncontent">{getRefferedNoteCount}</div>
      </Button>

      <Button
        className="landingPgTopBtn"
        style={{
          flex: '1',
          minWidth: '0',
          background: activeButton === "My Returned Notes" ? "#034ea1" : "",
          color: activeButton === "My Returned Notes" ? "#ffff" : "",
        }}
        onClick={() => handleButtonClick("My Returned Notes")}
      >
        My Returned Notes
        <div className="landingPgTopBtncontent">{returnNoteCount}</div>
      </Button>

      <Button
        className="landingPgTopBtn"
        style={{
          flex: '1',
          minWidth: '0',
          background: activeButton === "My Approved Notes" ? "#034ea1" : "",
          color: activeButton === "My Approved Notes" ? "#ffff" : "",
        }}
        onClick={() => handleButtonClick("My Approved Notes")}
      >
        My Approved Notes
        <div className="landingPgTopBtncontent">{approvedNotesCount}</div>
      </Button>

      <Button
        className="landingPgTopBtn"
        style={{
          flex: '1',
          minWidth: '0',
          background: activeButton === "My Upcoming Notes" ? "#034ea1" : "",
          color: activeButton === "My Upcoming Notes" ? "#ffff" : "",
        }}
        onClick={() => handleButtonClick("My Upcoming Notes")}
      >
        My Upcoming Notes
        <div className="landingPgTopBtncontent">{upcomingNotesCount}</div>
      </Button>

      {isSecretary && (
        <Button
          className="landingPgTopBtn"
          style={{
            flex: '1',
            minWidth: '0',
            background: activeButton === "ED/MD Notes" ? "#034ea1" : "",
            color: activeButton === "ED/MD Notes" ? "#ffff" : "",
          }}
          onClick={() => handleButtonClick("ED/MD Notes")}
        >
          ED/MD Notes
          <div className="landingPgTopBtncontent">{edMdNotesCount}</div>
        </Button>
      )}
    </div>
  );

  return (
    <div>
      <Navbar header="IB Smart Office - eNote" />
      {/* Task 319--showing the menus at first time--12/04--RK */}
      <Sidebar defaultOpenComponent={true} />
      {/* Added top navigation for menu --> 23-09 */}
      <MenuNavContainer isMenuPage />
      <div className="container largeScreen datagridpage">
        <div className="datagridPghdrBtns">
          <div className="row landingPgTopBtnRow">
            {isMobileView ? renderDropdown() : renderButtons()}
          </div>
        </div>

        {activeButton === "My Notes" ? (
          <ChartContainer apiOutput={apiOutputforChart} />
        ) : (
          <Dashboard apiOutput={apiOutput} selectedView={selectedView} enumObj={enumObj} />
        )}
      </div>
        {/* <ChatBot /> */}
      <div className="pgFooterContainer">
        <Footer />
      </div>
    </div>
  );
};
