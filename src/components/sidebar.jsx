import React, { useState, useEffect } from "react";
// Import Kendo Components
import {
  ExpansionPanel,
  ExpansionPanelContent,
} from "@progress/kendo-react-layout";
import { Reveal } from "@progress/kendo-react-animation";
// Import Kendo icon Components
import { chevronDownIcon, chevronUpIcon } from "@progress/kendo-svg-icons";
// Import External Libraries
import { Link } from "react-router-dom";
import { useMsal, useAccount } from "@azure/msal-react";
// Imported svg images
import list from "../../src/assets/list.svg";
import homeIcon from "../../src/assets/HomeIcon.svg"; // Imported Home icon for home button
// Import components
import { datagridlistitems } from "../pages/datagridlistitems";
import { getAccessToken } from "../App.js";
import { loginRequest } from "../config.js";
import { API_BASE_URL, API_ENDPOINTS, API_COMMON_HEADERS } from "../config.js";
import "../styles/responsiveDesign.css"

export const Sidebar = ({ defaultOpenComponent }) => {
  const [isFirstDivVisible, setFirstDivVisibility] = useState(false);
  const [isToggleClicked, setIsToggleClicked] = useState(false);
  const [expanded, setExpanded] = useState(false);
  const { instance, accounts } = useMsal();
  const [sidenavMenu, setSideNavMenu] = useState(datagridlistitems);
  const account = useAccount(accounts[0] || {});
  const [isATRSecretary, setIsATRSecretary] = useState(false); // New state for ATR Secretary

  const handleToggleHover = () => {
    setFirstDivVisibility(!isFirstDivVisible);
    setIsToggleClicked((prev) => !prev);
  };

  useEffect(() => {
    findDeptSuperAdmin();
  }, []);

  useEffect(() => {
    const handleOrientationChange = () => {
      const isLandscape = window.matchMedia("(orientation: landscape)").matches;
      if (isLandscape) {
        setFirstDivVisibility(false); // Close sidebar in landscape mode
        setIsToggleClicked(false);
      }
    };

    // Initial check
    handleOrientationChange();

    // Add event listener for orientation changes
    window.addEventListener("resize", handleOrientationChange);

    // Cleanup
    return () => window.removeEventListener("resize", handleOrientationChange);
  }, []);

  useEffect(() => {
    if (defaultOpenComponent) {
      setFirstDivVisibility(!isFirstDivVisible);
      setIsToggleClicked(true);
    }
  }, [defaultOpenComponent]);

  const findIsATRSecretary = async () => {
    let isSecretaryAtr = false;
  
    try {
      const accessToken = await getAccessToken({
        ...loginRequest,
        account,
      }, instance);
  
      const response = await fetch(
        `${API_BASE_URL}${API_ENDPOINTS.eNote_findATRSecretaries}`,
        {
          method: "POST",
          body: JSON.stringify({ SecretaryEmail: accounts[0].username }),
          headers: {
            ...API_COMMON_HEADERS,
            Authorization: `Bearer ${accessToken}`,
          },
        }
      );
  
      const data = await response.json();
  
      isSecretaryAtr = data === true; // Ensure this matches API response
      setIsATRSecretary(isSecretaryAtr);
  
      // Optionally return the value for further conditional logic
      return isSecretaryAtr;
  
    } catch (error) {
      console.error("Error fetching ATR Secretary status:", error);
      setIsATRSecretary(false);
      return false;
    }
  };
  

  const findDeptSuperAdmin = async () => {
    const accessToken = await getAccessToken({
      ...loginRequest,
      account,
    }, instance);
  
    try {
      const response = await fetch(`${API_BASE_URL}${API_ENDPOINTS.Find_Super_Dept_Admin(accounts[0]?.username)}`, {
        method: "GET",
        headers: { ...API_COMMON_HEADERS, Authorization: `Bearer ${accessToken}` },
      });
  
      const data = await response.json();
  
      const secretaryATR = await findIsATRSecretary();
      
      if (data.isSuperAdmin) {
        setSideNavMenu(datagridlistitems);
      } else if (data.isDepartmentAdmin || secretaryATR) {
        const validatedMenu = datagridlistitems.filter(item => item.id !== "Administration");
        const validateAtr = validatedMenu.map(item =>
          item.id === "ATR Views"
            ? { ...item, items: item.items.filter(subItem => subItem !== "All ATR" || secretaryATR) }
            : item
        );
        setSideNavMenu(validateAtr);
      } else {
        const validatedMenu = datagridlistitems.filter(item => item.id !== "Administration" && item.id !== "Department Admin");
        const validateAtr = validatedMenu.map(item =>
          item.id === "ATR Views"
            ? { ...item, items: item.items.filter(subItem => subItem !== "All ATR" || secretaryATR) }
            : item.id === "Reports"
              ? { ...item, items: item.items.filter(subItem => subItem === "Dynamic Search" || secretaryATR) }
              : item
        );
        setSideNavMenu(validateAtr);
      }
    } catch (error) {
      console.error("Error fetching department admin:", error);
    }
  };
  

  const getLinkForItem = (item) => {
    if (item === sidenavMenu[1].items[0]) {
      return "/eNoteForm";
    } else if (item === sidenavMenu[sidenavMenu.findIndex((obj) => obj.id === "Views")].items[0]) {
      return `/views/${sidenavMenu[sidenavMenu.findIndex((obj) => obj.id === "Views")].items[0]}`;
    } else if (item === sidenavMenu[sidenavMenu.findIndex((obj) => obj.id === "Views")].items[1]) {
      return `/views/${sidenavMenu[sidenavMenu.findIndex((obj) => obj.id === "Views")].items[1]}`;
    } else if (item === sidenavMenu[sidenavMenu.findIndex((obj) => obj.id === "Views")].items[2]) {
      return `/views/${sidenavMenu[sidenavMenu.findIndex((obj) => obj.id === "Views")].items[2]}`;
    } else if (item === sidenavMenu[sidenavMenu.findIndex((obj) => obj.id === "Views")].items[3]) {
      return `/views/${sidenavMenu[sidenavMenu.findIndex((obj) => obj.id === "Views")].items[3]}`;
    } else if (item === sidenavMenu[sidenavMenu.findIndex((obj) => obj.id === "Views")].items[4]) {
      return `/views/${sidenavMenu[sidenavMenu.findIndex((obj) => obj.id === "Views")].items[4]}`;
    } else if (item === sidenavMenu[sidenavMenu.findIndex((obj) => obj.id === "Views")].items[5]) {
      return `/views/${sidenavMenu[sidenavMenu.findIndex((obj) => obj.id === "Views")].items[5]}`;
    } else if (item === sidenavMenu[sidenavMenu.findIndex((obj) => obj.id === "ATR Views")].items[0]) {
      return `/atrviews/${sidenavMenu[sidenavMenu.findIndex((obj) => obj.id === "ATR Views")].items[0]}`;
    } else if (item === sidenavMenu[sidenavMenu.findIndex((obj) => obj.id === "ATR Views")].items[1]) {
      return `/atrviews/${sidenavMenu[sidenavMenu.findIndex((obj) => obj.id === "ATR Views")].items[1]}`;
    } else if (item === sidenavMenu[sidenavMenu.findIndex((obj) => obj.id === "ATR Views")].items[2]) {
      return `/atrviews/${sidenavMenu[sidenavMenu.findIndex((obj) => obj.id === "ATR Views")].items[2]}`;
    } else if (item === sidenavMenu[sidenavMenu.findIndex((obj) => obj.id === "ATR Views")].items[3]) {
      return `/atrviews/${sidenavMenu[sidenavMenu.findIndex((obj) => obj.id === "ATR Views")].items[3]}`;
    } else if (item === sidenavMenu[sidenavMenu.findIndex((obj) => obj.id === "ATR Views")].items[4]) {
      return `/atrviews/${sidenavMenu[sidenavMenu.findIndex((obj) => obj.id === "ATR Views")].items[4]}`;
    } else if (item === sidenavMenu[sidenavMenu.findIndex((obj) => obj.id === "Reports")].items[0]) {
      return "/searchnotes";
    } else if (item === sidenavMenu[sidenavMenu.findIndex((obj) => obj.id === "Reports")].items[1]) {
      return "/dashBoardReports";
    }
    else {
      return "";
    }
  };

  return (
    <div className="_mobileResponsive_mobile">
      <div className="homeButtonContainer">
        <Link to="/datagridpage" className="homeButton">
          <img src={homeIcon} alt="Home Icon" className="homeIcon" />
          <span className="homeTooltip">eNote Home</span>
        </Link>
      </div>
      {isFirstDivVisible && (
        <div className="sideBarContainer">
          {sidenavMenu.map((item, index) => (
            <ExpansionPanel
              key={index}
              style={{ margin: "10px" }}
              title={
                <div style={{ display: "flex" }} className="mobile_menu">
                  <div
                    style={{ width: "18px", marginRight: "6px" }}
                    dangerouslySetInnerHTML={{ __html: item.image }}
                  />
                  {item.id}
                  {index !== 0 && (
                    <div style={{ marginLeft: "auto" }}></div>
                  )}
                </div>
              }
              collapseSVGIcon={item.id === "Home" || item.id === "Create eNote" || item.id === "Create/Reset Passcode" ? (
                <svg aria-hidden="true" focusable="false" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512"></svg>
              ) : (
                chevronUpIcon
              )}
              expandSVGIcon={item.id === "Home" || item.id === "Create eNote" || item.id === "Create/Reset Passcode" ? (
                <svg aria-hidden="true" focusable="false" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512"></svg>
              ) : (
                chevronDownIcon
              )}
              expanded={expanded === item.id}
              tabIndex={0}
              expandIcon=""
              onAction={(event) => {
                if (item.id.includes("Create eNote")) {
                  window.location.href = "/enoteform/new";
                }
                else if (item.id.includes("Home")) {
                  window.location.href = "/datagridpage";
                }
                else if (item.id.includes("Create/Reset Passcode")) {
                  window.location.href = "/enotepasscode"
                }
                else {
                  setExpanded(event.expanded ? "" : item.id);
                }
              }}
              className={index === 0 ? "hide-default-icon" : ""}
            >
              <Reveal>
                {expanded === item.id && (
                  <ExpansionPanelContent>
                    <ul style={{ paddingLeft: "0px", marginBottom: "0px" }}>
                      {item.items.map((listItem, listItemIndex) =>
                        index !== 0 && (
                          <Link
                            to={getLinkForItem(listItem)}
                            style={{ textDecoration: "none" }}
                            onClick={handleToggleHover}
                            key={listItemIndex}
                          >
                            <li
                              className="list-item"
                              style={{ listStyleType: "none" }}
                            >
                              {listItem}
                            </li>
                          </Link>
                        )
                      )}
                    </ul>
                  </ExpansionPanelContent>
                )}
              </Reveal>
            </ExpansionPanel>
          ))}
        </div>
      )}
      <div
        className="sideBarClose"
        style={{
          paddingLeft: isToggleClicked ? "20%" : "30px",
          paddingRight: isToggleClicked ? "auto" : "auto",
        }}
      >
        {!isToggleClicked && (
          <img
            src={list}
            className="cursor homeLogo"
            alt="HomeIcon"
            onClick={handleToggleHover}
          />
        )}
        {isToggleClicked && (
          <span className="k-icon k-font-icon k-i-close cursor sideBarmenuCloseIc" onClick={handleToggleHover}></span>
        )}
      </div>
    </div>
  );
};
