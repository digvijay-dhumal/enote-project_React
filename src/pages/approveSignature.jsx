import React, { useEffect, useState } from "react";
import { Grid, GridColumn } from "@progress/kendo-react-grid";
import { process } from "@progress/kendo-data-query";
import { Button } from "@progress/kendo-react-buttons";
import { Dialog, DialogActionsBar } from "@progress/kendo-react-dialogs";
import { useMsal, useAccount } from "@azure/msal-react";
import CryptoJS from "crypto-js";

import Navbar from "../components/navbar";
import Footer from "../components/footer";
import { Sidebar } from "../components/sidebar";
import { MenuNavContainer } from "../components/menu";
import view from "../assets/view.png"
import hide from "../assets/hide.png"
import { getAccessToken } from "../App";
import {
    API_BASE_URL,
    API_COMMON_HEADERS,
    loginRequest,
     API_ENDPOINTS
} from "../config";

import "../styles/approve-signature.css";
import "../styles/passcode.css";


export const ApproveSignature = () => {
    const { instance, accounts } = useMsal();
    const account = useAccount(accounts[0]);

    const [data, setData] = useState([]);
    const [loading, setLoading] = useState(false);
    const [dataState, setDataState] = useState({
        skip: 0,
        take: 10
    });

    // State for passcode verification
    const [passcodeVerification, setPasscodeVerification] = useState(false);
    const [passcode, setPasscode] = useState("");
    const [isPasscodeVisible, setIsPasscodeVisible] = useState(false);
    const [error, setError] = useState("");
    const [actionBtn, setActionBtn] = useState(""); // Store current action type
    const [selectedItem, setSelectedItem] = useState(null); // Store selected data item

    // State for confirmation dialog
    const [dialogOpen, setDialogOpen] = useState(false);
    const [confirmDialogObj, setConfirmDialogObj] = useState({
        Confirmtext: "",
        Description: ""
    });

    useEffect(() => {
        if (account && accounts.length > 0) {
            fetchPendingApprovals();
        }
    }, [account]);


     const fetchBase64Image = async (blobPath) => {
        const accessToken = await getAccessToken(
            { ...loginRequest, account },
            instance
        );

        const response = await fetch(
            `${API_BASE_URL}/api/ENote/GetBase64`,
            {
                method: "POST",
                headers: {
                    ...API_COMMON_HEADERS,
                    Authorization: `Bearer ${accessToken}`,
                    "Content-Type": "application/json"
                },
                body: JSON.stringify({
                    supportingDocumentPath: blobPath
                })
            }
        );

        if (!response.ok) {
            throw new Error("Failed to fetch base64 image");
        }

        const base64 = await response.text();
        return `data:image/png;base64,${base64}`;
    };

    /* ----------------------------------
       FETCH PENDING SIGNATURES
    -----------------------------------*/

      const fetchPendingApprovals = async () => {
        setLoading(true);
        try {
            const accessToken = await getAccessToken(
                { ...loginRequest, account },
                instance
            );

            const approverMail = accounts[0].username;

            const response = await fetch(
                `${API_BASE_URL}/api/Admin/GetApproverPendingList?approverMail=${encodeURIComponent(
                    approverMail
                )}`,
                {
                    method: "GET",
                    headers: {
                        ...API_COMMON_HEADERS,
                        Authorization: `Bearer ${accessToken}`
                    }
                }
            );

            if (!response.ok) {
                throw new Error(`API failed ${response.status}`);
            }

            const result = await response.json();

            const mappedData = await Promise.all(
                result.map(async (item, index) => {
                    let signUrl = null;
                    try {
                        signUrl = await fetchBase64Image(item.pendingBlobPath);
                    } catch (err) {
                        console.error("Image load error", err);
                    }

                    return {
                        id: index,
                        userPrincipalName: item.empName,
                        srNo: item.empId,
                        signUrl
                    };
                })
            );

            setData(mappedData);
        } catch (err) {
            console.error(err);
            setData([]);
        } finally {
            setLoading(false);
        }
    };
    
    // const fetchPendingApprovals = async () => {
    //     setLoading(true);
    //     try {
    //         const accessToken = await getAccessToken(
    //             { ...loginRequest, account },
    //             instance
    //         );

    //         const approverMail = accounts[0].username;

    //         const url = `${API_BASE_URL}/api/Admin/GetApproverPendingList?approverMail=${encodeURIComponent(
    //             approverMail
    //         )}`;

    //         const response = await fetch(url, {
    //             method: "GET",
    //             headers: {
    //                 ...API_COMMON_HEADERS,
    //                 Authorization: `Bearer ${accessToken}`
    //             }
    //         });

    //         if (!response.ok) {
    //             throw new Error(`API failed with ${response.status}`);
    //         }

    //         const result = await response.json();

    //         /* MAP API RESPONSE */
    //         const mappedData = result.map((item, index) => ({
    //             id: item.id ?? index,
    //             userPrincipalName: item.empName,
    //             srNo: item.empId,
    //             signUrl: item.pendingBlobPath // Azure Blob URL
    //         }));

    //         setData(mappedData);
    //     } catch (error) {
    //         console.error("Fetch Error:", error);
    //         setData([]);
    //     } finally {
    //         setLoading(false);
    //     }
    // };

    /* ----------------------------------
       PASSCODE VERIFICATION HANDLERS
    -----------------------------------*/
    const handlePasscodeChange = (e) => {
        const value = e.target.value;
        // Allow only alphanumeric characters
        if (/^[a-zA-Z0-9]*$/.test(value)) {
            setPasscode(value);
            if (error) setError("");
        }
    };

    const handleKeyDown = (e) => {
        if (e.key === "Enter" || e.key === " ") {
            setIsPasscodeVisible(true);
        }
    };

    const handleKeyUp = (e) => {
        if (e.key === "Enter" || e.key === " ") {
            setIsPasscodeVisible(false);
        }
    };

    const handlePasscodeClose = () => {
        setPasscodeVerification(false);
        setPasscode("");
        setIsPasscodeVisible(false);
        setError("");
        setSelectedItem(null);
        setActionBtn("");
    };

    /* ----------------------------------
       PASSCODE VERIFICATION FUNCTION
    -----------------------------------*/
    const passcodeVerificationFunction = async () => {
        try {
            const accessToken = await getAccessToken(
                { ...loginRequest, account },
                instance
            );

            const secretKey = "SHA256";
            const hashedPasscode = CryptoJS.SHA256(passcode, secretKey).toString(CryptoJS.enc.Hex);

            const params = {
                Passcode: hashedPasscode,
                UserMail: accounts[0].username
            };

            // FIX: Use the correct API endpoint for passcode verification
            const response = await fetch(
                  `${API_BASE_URL}${API_ENDPOINTS.eNote_VerifyPasscode}`, // Fixed endpoint
                {
                    method: "POST",
                    body: JSON.stringify(params),
                    headers: {
                        ...API_COMMON_HEADERS,
                        Authorization: `Bearer ${accessToken}`,
                        'Content-Type': 'application/json'
                    },
                }
            );

            if (!response.ok) {
                throw new Error(`Passcode verification failed: ${response.status}`);
            }

            const data = await response.text();

            if (data === "Passcode verification successful") {
                setPasscodeVerification(false);
                setPasscode("");
                
                // Show confirmation dialog based on action
                if (actionBtn === "Approve") {
                    setConfirmDialogObj({
                        Confirmtext: "Are you sure you want to approve this signature?",
                        Description: "Please verify the signature details before approving."
                    });
                } else if (actionBtn === "Reject") {
                    setConfirmDialogObj({
                        Confirmtext: "Are you sure you want to reject this signature?",
                        Description: "Please verify the signature details before rejecting."
                    });
                }
                
                setDialogOpen(true);
            } else {
                setError('Incorrect passcode entered. Please try again.');
            }
        } catch (err) {
            console.error("Passcode verification error:", err);
            setError('Verification failed. Please try again.');
        }
    };

    /* ----------------------------------
       SEPARATE API CALLS FOR APPROVE/REJECT
    -----------------------------------*/
    const approveSignature = async () => {
        try {
            const accessToken = await getAccessToken(
                { ...loginRequest, account },
                instance
            );

            const payload = {
                empId: selectedItem.srNo,
                approverMail: accounts[0].username
            };

            const response = await fetch(
                `${API_BASE_URL}/api/Admin/ApproveSignature`,
                {
                    method: "POST",
                    headers: {
                        ...API_COMMON_HEADERS,
                        Authorization: `Bearer ${accessToken}`,
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify(payload)
                }
            );

            if (!response.ok) {
                const errorText = await response.text();
                throw new Error(`Approve failed: ${response.status} - ${errorText}`);
            }

            // Refresh the data
            fetchPendingApprovals();
            
            // Close dialog
            setDialogOpen(false);
            setSelectedItem(null);
            setActionBtn("");
            
            // Show success message (optional)
            console.log("Signature approved successfully");
            
        } catch (error) {
            console.error("Approve Error:", error);
            setError(`Approve failed: ${error.message}`);
            // Keep dialog open to show error
        }
    };

    const rejectSignature = async () => {
        try {
            const accessToken = await getAccessToken(
                { ...loginRequest, account },
                instance
            );

            const payload = {
                empId: selectedItem.srNo,
                approverMail: accounts[0].username
            };

            const response = await fetch(
                `${API_BASE_URL}/api/Admin/RejectSignature`,
                {
                    method: "POST",
                    headers: {
                        ...API_COMMON_HEADERS,
                        Authorization: `Bearer ${accessToken}`,
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify(payload)
                }
            );

            if (!response.ok) {
                const errorText = await response.text();
                throw new Error(`Reject failed: ${response.status} - ${errorText}`);
            }

            // Refresh the data
            fetchPendingApprovals();
            
            // Close dialog
            setDialogOpen(false);
            setSelectedItem(null);
            setActionBtn("");
            
            // Show success message (optional)
            console.log("Signature rejected successfully");
            
        } catch (error) {
            console.error("Reject Error:", error);
            setError(`Reject failed: ${error.message}`);
            // Keep dialog open to show error
        }
    };

    /* ----------------------------------
       ACTION HANDLERS
    -----------------------------------*/
    const handleApproveClick = (dataItem) => {
        setSelectedItem(dataItem);
        setActionBtn("Approve");
        setPasscodeVerification(true);
    };

    const handleRejectClick = (dataItem) => {
        setSelectedItem(dataItem);
        setActionBtn("Reject");
        setPasscodeVerification(true);
    };

    /* ----------------------------------
       CONFIRMATION HANDLERS
    -----------------------------------*/
    const toggleDialogForCancel = () => {
        setDialogOpen(false);
        setSelectedItem(null);
        setActionBtn("");
        setError("");
    };

    const onConfirmation = async () => {
        setError(""); // Clear any previous errors
        
        try {
            if (actionBtn === "Approve") {
                await approveSignature();
            } else if (actionBtn === "Reject") {
                await rejectSignature();
            }
        } catch (error) {
            console.error("Confirmation Error:", error);
            setError(`Operation failed: ${error.message}`);
        }
    };

    /* ----------------------------------
       UI COMPONENTS
    -----------------------------------*/
    // const SignCell = (props) => (
    //     <td className="center-cell">
    //         {props.dataItem.signUrl ? (
    //             <img
    //                 src={props.dataItem.signUrl}
    //                 alt="Signature"
    //                 className="sign-preview"
    //             />
    //         ) : (
    //             <span>No Signature</span>
    //         )}
    //     </td>
    // );


     const SignCell = (props) => (
        <td className="center-cell">
            {props.dataItem.signUrl ? (
                <img
                    src={props.dataItem.signUrl}
                    alt="Signature"
                    className="sign-preview"
                />
            ) : (
                "No Signature"
            )}
        </td>
    );
    
    const ActionCell = (props) => (
        <td className="center-cell">
            <Button
                themeColor="success"
                size="small"
                className="me-2"
                onClick={() => handleApproveClick(props.dataItem)}
            >
                Approve
            </Button>

            <Button
                themeColor="error"
                size="small"
                onClick={() => handleRejectClick(props.dataItem)}
            >
                Reject
            </Button>
        </td>
    );

    const CustomConfirmationTitleBar = () => (
        <div className="dialog-title-bar">
            <span className="k-icon k-font-icon k-i-warning"></span>
            <span>Confirmation</span>
        </div>
    );

    /* ----------------------------------
       UI
    -----------------------------------*/
    return (
        <div className="app-layout">
            <Navbar header="IB Smart Office - eNote" />

            <div className="body-layout">
                <Sidebar />

                <div className="content-layout">
                    <MenuNavContainer isMenuPage />

                    <div className="container largeScreen text-center">
                        <div className="SectionHeads row mobileSectionHeads">
                            Approve Signature
                        </div>
                    </div>
                    
                    <div className="grid-wrapper">
                        <Grid
                            data={process(data, dataState)}
                            pageable={{ pageSizes: [10, 20, 50] }}
                            sortable
                            loading={loading}
                            onDataStateChange={(e) =>
                                setDataState(e.dataState)
                            }
                            className="exact-grid"
                        >
                            <GridColumn
                                field="userPrincipalName"
                                title="User Principal Name"
                                width={350}
                            />

                            <GridColumn
                                field="srNo"
                                title="SR No"
                                width={150}
                            />

                            <GridColumn
                                title="Signature"
                                cell={SignCell}
                                width={250}
                            />

                            <GridColumn
                                title="Action"
                                cell={ActionCell}
                                width={300}
                            />
                        </Grid>
                    </div>
                </div>
            </div>

            <Footer />

            {/* Passcode Verification Dialog */}
            {passcodeVerification && (
                <Dialog 
                    title="Passcode Verification" 
                    onClose={handlePasscodeClose}
                    className="passcode-dialog"
                >
                    <form className="form-container dialogcontent_">
                        <label>
                            Enter your passcode for verification:
                            <div className="passcode-input-wrapper">
                                <input
                                    type={isPasscodeVisible ? 'text' : 'password'}
                                    value={passcode}
                                    onChange={handlePasscodeChange}
                                    className="passcode-input"
                                    maxLength="6"
                                    pattern="[a-zA-Z0-9]*"
                                    title="Please enter 6-characters, Combination of Alphabets and Numbers"
                                />
                                <span
                                    className="eye-icon-view"
                                    tabIndex="0"
                                    onMouseDown={() => setIsPasscodeVisible(true)}
                                    onMouseUp={() => setIsPasscodeVisible(false)}
                                    onKeyDown={(e) => handleKeyDown(e)}
                                    onKeyUp={(e) => handleKeyUp(e)}
                                    onTouchStart={() => setIsPasscodeVisible(true)}
                                    onTouchEnd={() => setIsPasscodeVisible(false)}
                                >
                                    {isPasscodeVisible ? (
                                        <img alt="view" src={view} />
                                    ) : (
                                        <img alt="hide" src={view} />
                                    )}
                                </span>
                            </div>
                        </label>
                        <div>
                            {error && <span className='error'>{error}</span>}
                        </div>
                    </form>
                    <DialogActionsBar>
                        <Button
                            className="formBtnColor"
                            onClick={passcodeVerificationFunction}
                            disabled={passcode.length !== 6}
                        >
                            <span className="k-icon k-font-icon k-i-launch cursor allIconsforPrimary-btn"></span>
                            Verify
                        </Button>
                        <Button onClick={handlePasscodeClose}>
                            <span className="k-icon k-font-icon k-i-close-circle cursor allIconsforPrimary-btn"></span>
                            Cancel
                        </Button>
                    </DialogActionsBar>
                </Dialog>
            )}

            {/* Confirmation Dialog */}
            {dialogOpen && (
                <Dialog
                    title={<CustomConfirmationTitleBar />}
                    onClose={toggleDialogForCancel}
                    className="confirmation-dialog"
                >
                    <p className="dialogcontent_">
                        {confirmDialogObj.Confirmtext}
                    </p>
                    <p className="dialogcontent_">
                        {confirmDialogObj.Description}
                    </p>
                    
                    {/* Show error message if any */}
                    {error && (
                        <div className="error-message mb-3">
                            <span className="k-icon k-font-icon k-i-warning"></span>
                            {error}
                        </div>
                    )}

                    <DialogActionsBar>
                        <Button 
                            className="formBtnColor" 
                            onClick={onConfirmation}
                        >
                            <span className="k-icon k-font-icon k-i-checkmark-circle cursor allIconsforPrimary-btn"></span>
                            Confirm
                        </Button>

                        <Button onClick={toggleDialogForCancel}>
                            <span className="k-icon k-font-icon k-i-close-circle cursor allIconsforPrimary-btn"></span>
                            Cancel
                        </Button>
                    </DialogActionsBar>
                </Dialog>
            )}
        </div>
    );
};