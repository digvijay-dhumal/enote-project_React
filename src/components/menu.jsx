import React, { useState, useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import { Menu, MenuItem } from '@progress/kendo-react-layout';
import { useMsal, useAccount } from "@azure/msal-react";
import { datagridlistitems } from '../pages/datagridlistitems';
import { getAccessToken } from "../App.js";
import { API_BASE_URL, API_ENDPOINTS, API_COMMON_HEADERS, loginRequest, IB_eArchivalDomain_URL } from "../config.js";
import "../styles/navbar.css";

export const MenuNavContainer = (props) => {
    const navigate = useNavigate();
    const [sidenavMenu, setSideNavMenu] = useState(datagridlistitems);
    const { instance, accounts } = useMsal();
    const account = useAccount(accounts[0] || {});
    const [isATRSecretary, setIsATRSecretary] = useState(false); // New state for ATR Secretary

    useEffect(() => {
        findDeptSuperAdmin();
    }, []);

    const findIsATRSecratary = async () => {
        let isSecretaryAtr = false;
        const accessToken = await getAccessToken({
            ...loginRequest,
            account,
        }, instance);

        await fetch(
            `${API_BASE_URL}${API_ENDPOINTS.eNote_findATRSecretaries}`,
            {
                method: "POST",
                body: JSON.stringify({ SecretaryEmail: accounts[0].username }),
                headers: { ...API_COMMON_HEADERS, Authorization: `Bearer ${accessToken}` },
            }
        )
            .then((res) => res.json())
            .then((data) => {
                isSecretaryAtr = data;
                setIsATRSecretary(data);
            })
            .catch((error) => {
                console.error("Error fetching data:", error);
            });
        return isSecretaryAtr;
    };

    const findDeptSuperAdmin = async () => {

        const accessToken = await getAccessToken({
            ...loginRequest,
            account
        }, instance);

        try {
            const response = await fetch(`${API_BASE_URL}${API_ENDPOINTS.Find_Super_Dept_Admin(accounts[0]?.username)}`, {
                method: "GET",
                headers: { ...API_COMMON_HEADERS, Authorization: `Bearer ${accessToken}` },
            });

            const data = await response.json();

            const secretaryATR = await findIsATRSecratary();
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
            }
            else {
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

    const onSelect = (event) => {
        const selectedItem = event.item.text;

        if (selectedItem === "Create eNote") {
            window.location.href = "/enoteform/new";
            return;
        }
        if (selectedItem === "Home") {
            window.location.href = "/datagridpage";
            return;
        }
        if (selectedItem === "Create/Reset Passcode") {
            window.location.href = "/enotepasscode";
            return;
        }
         if (selectedItem === "Upload Signature") {
        navigate("/uploadSignature"); // no reload
        return;
    }
    if (selectedItem === "Approve Signature") {
        navigate("/approveSignature"); // no reload
        return;
    }

        // Handle eArchival menu items directly
        if (selectedItem === "eNote Search" || selectedItem === "eNote ATR Search") {
            let base = IB_eArchivalDomain_URL;
            if (!base || base === 'undefined') {
                alert('Archival domain URL is not set in .env!');
                return;
            }
            // Remove trailing slash if present
            if (base.endsWith('/')) base = base.slice(0, -1);
            if (selectedItem === "eNote Search") {
                window.location.href = `${base}/eNoteSearch?from=enote`;
                return;
            }
            if (selectedItem === "eNote ATR Search") {
                window.location.href = `${base}/enoteATRSearch?from=enote`;
                return;
            }
        }

        if (event.item.data.route) {
            navigate(event.item.data.route);
        }
    };

    const getLinkForItem = (item, parentId) => {
        if (parentId === "Views") {
            return `/views/${encodeURIComponent(item)}`;
        } else if (parentId === "ATR Views") {
            return `/atrviews/${encodeURIComponent(item)}`;
        } else if (parentId === "Reports") {
            if (item === "Dynamic Search") return "/searchnotes";
            if (item === "DashBoard Reports") return "/dashBoardReports";
        }
        return "";
    };

    const generateMenuItems = () => {
        return sidenavMenu.map((menuItem) => (
            <MenuItem
                key={menuItem.id}
                text={menuItem.id}
                data={{ route: "" }}
            >
                {menuItem.items && menuItem.items.map((subItem, index) => {
                    // Conditional rendering for "ATR Closure"
                    if (subItem === "ATR Closure" && !isATRSecretary) {
                        return null; // Don't render if not an ATR Secretary
                    }
                    return (
                        <MenuItem
                            key={index}
                            text={subItem}
                            data={{ route: getLinkForItem(subItem, menuItem.id) }}
                        />
                    );
                })}
            </MenuItem>
        ));
    };

    return (
        <div className={`container menu-container non_mobileResponsive_menu ${props.isViewPage ? 'largeScreen' : ''} ${props.isMenuPage ? 'largeScreen' : ''}`}>
            <Menu onSelect={onSelect}>
                {generateMenuItems()}
            </Menu>
        </div>
    );
};
