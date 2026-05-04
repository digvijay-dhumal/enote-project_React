import React, { useState, useRef } from 'react';
// import kendo components
import { Button } from "@progress/kendo-react-buttons";
import { Dialog, DialogActionsBar } from "@progress/kendo-react-dialogs";
// import external components
import CryptoJS from 'crypto-js';
import { useNavigate } from 'react-router-dom';
import { useMsal, useAccount } from "@azure/msal-react";
// import internal components and pages
import Navbar from "../components/navbar";
import Footer from "../components/footer";
import { Sidebar } from "../components/sidebar";
import { MenuNavContainer } from "../components/menu";
import { PageLoader } from "../components/pageLoader";
import { getAccessToken, useTabContext } from "../App";
import { API_COMMON_HEADERS, API_BASE_URL, API_ENDPOINTS, loginRequest } from "../config";
// import png
import view from "../assets/view.png";
import hide from "../assets/hide.png";
// import css styles
import "../styles/passcode.css";

export const ENotePasscode = () => {
    const { passcodenaviagate } = useTabContext();
    const navigate = useNavigate()
    const [passcodeFormType, setPasscodeFormType] = useState(passcodenaviagate)
    const [passcode, setPasscode] = useState('');
    const [confimPasscode, setConfimPasscode] = useState('');
    const [otp, setOtp] = useState(new Array(6).fill(''));
    const [passcodeSent, setPasscodeSent] = useState(false);
    const [passcodeVerified, setPasscodeVerified] = useState(false);
    const [error, setError] = useState('');
    const [timer, setTimer] = useState(180); // 3 minutes in seconds
    const [timerExpired, setTimerExpired] = useState(false);
    const [timerId, setTimerId] = useState(null);
    const inputsRef = useRef([]);
    const [verifyContentHide, setVerifyContentHide] = useState(true);
    const [saveButtonVisible, setSaveButtonVisible] = useState(false);
    const [showSuccessPopup, setShowSuccessPopup] = useState(false);
    const [isPasscodeVisible, setIsPasscodeVisible] = useState(false);
    const [isConfirmPasscodeVisible, setIsConfirmPasscodeVisible] = useState(false);
    const [showErrorPopup, setShowErrorPopup] = useState(false);
    const { accounts, instance } = useMsal();
    const account = useAccount(accounts[0] || {});
    const [encryptedPasscode, setEncryptedPasscode] = useState('');
    const [passcodeSaved, setPasscodeSaved] = useState(false);
    const [isLoading, setIsLoading] = useState(false);

    // Function to generate a random 6-digit OTP
    const generateOTP = () => {
        return Math.floor(100000 + Math.random() * 900000);
    };

    // Function to handle passcode submission
    const handleAddPasscode = async (event) => {
        setIsLoading(true);
        event.preventDefault();
        const regex = /^(?=.*[A-Za-z])(?=.*\d)[A-Za-z\d]+$/;
        if (passcode.length < 6 || confimPasscode.length < 6) {
            setError('Passcode should contain 6-characters.');
        }
        else if (!passcode || !confimPasscode) {
            setError('Passcode and confirm passcode should not be empty.');
            return;
        } else if (!regex.test(passcode) || !regex.test(confimPasscode)) {
            setError('Passcode should contain both Alphabets and Numbers.');
            return;
        }
        else if (passcode !== confimPasscode) {
            setError('Passcode and confirm passcode should match.');
            return;
        } else {

            const otp = generateOTP(); // Generate OTP here

            try {
                const accessToken = await getAccessToken({ ...loginRequest, account }, instance);
                // Replace with your own secret key
                const secretKey = CryptoJS.enc.Utf8.parse("your-256-bit-secret-key-12345678");
                const iv = CryptoJS.enc.Utf8.parse("your-128-bit-iv1");
                // Encrypt the passcode with the secret key and IV
                const encryptedPasscode = CryptoJS.AES.encrypt(otp.toString(), secretKey, { iv: iv }).toString();

                setEncryptedPasscode(encryptedPasscode);

                const params = {
                    OTP: encryptedPasscode,
                    UserMail: accounts[0].username
                };

                const response = await fetch(
                    `${API_BASE_URL}${API_ENDPOINTS.eNote_SendOTP}`,
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

                const data = await response.text();

                if (data === "OTP SENT") {
                    setIsLoading(false);
                    setPasscodeSent(true);
                    startTimer(); // Start the timer after successfully sending the OTP 
                } else {
                    setError('Incorrect OTP entered please try again.');
                }
            } catch (err) {
                console.error(err);
            }

        }
    };

    // Function to start the 3-minute timer
    const startTimer = () => {
        const id = setInterval(() => {
            setTimer((prev) => {
                if (prev <= 1) {
                    clearInterval(id);
                    setTimerExpired(true);
                    return 0;
                }
                return prev - 1;
            });
        }, 1000); // Update timer every second
        setTimerId(id); // Store the interval ID in the state
    };

    // Function to verify the passcode entered by the user
    const handleVerifyPasscode = async () => {
        clearInterval(timerId); // Stop the timer
        setVerifyContentHide(false); // Hide the timer text when verify button is clicked
        const enteredOtp = otp.join('');
        if (enteredOtp.length !== 6) {
            setError('Please enter the 6-digit verification code.');
            return;
        }

        // Replace with your own secret key
        const secretKey = CryptoJS.enc.Utf8.parse("your-256-bit-secret-key-12345678");
        const iv = CryptoJS.enc.Utf8.parse("your-128-bit-iv1");

        // Decrypt the passcode stored in state
        const decryptedPasscode = CryptoJS.AES.decrypt(encryptedPasscode, secretKey, { iv: iv }).toString(CryptoJS.enc.Utf8);

        // Compare decrypted passcode with user entered OTP
        if (enteredOtp === decryptedPasscode) {
            setPasscodeVerified(true);
            setSaveButtonVisible(true); // Show the save button when passcode is verified
        } else {
            setShowErrorPopup(true);
            setError('Incorrect passcode entered please try again.');
        }
    };

    // Function to save the passcode to local storage
    const handleSavePasscode = async () => {
        setShowSuccessPopup(true); // Show success popup
        const secretKey = 'SHA256'; // Replace with your own secret key

        // Use SHA-256 to hash the passcode with the secret key
        const hash = CryptoJS.SHA256(passcode, secretKey);
        const encryptedPasscode = hash.toString(CryptoJS.enc.Hex);

        try {
            const accessToken = await getAccessToken({ ...loginRequest, account }, instance);
            const params = {
                Passcode: encryptedPasscode,
                UserMail: accounts[0].username
            };

            const response = await fetch(
                `${API_BASE_URL}${API_ENDPOINTS.eNote_AddPasscode}`,
                {
                    method: "POST",
                    body: JSON.stringify(params),
                    headers: { ...API_COMMON_HEADERS, Authorization: `Bearer ${accessToken}` }
                }
            );

            const data = await response.json();
            setPasscodeSaved(true);
        } catch (err) {
            console.error(err);
        }
    };

    // Function to handle closing the success popup
    const handleCloseSuccessPopup = () => {
        setShowSuccessPopup(false);
        setSaveButtonVisible(false);
        setPasscode('');
        setConfimPasscode("")
        setOtp(new Array(6).fill(''));
        setPasscodeSent(false);
        setPasscodeVerified(false);
        setError('');
        setTimer(180); // Reset timer to 3 minutes
        setTimerExpired(false);
        setVerifyContentHide(true);
        if (passcodeFormType !== "New") {
            navigate(-1)
        }
    };

    // Function to resend the passcode
    const handleResendPasscode = async (event) => {
        clearInterval(timerId); // Clear any existing interval
        setPasscode('');
        setConfimPasscode("");
        setOtp(new Array(6).fill(''));
        setTimerExpired(false);
        setTimer(180); // Reset timer to 3 minutes
        handleAddPasscode(event);
    };

    // Handle input change in OTP fields
    const handleOtpChange = (index, value) => {
        if (/[^a-zA-Z0-9]/.test(value)) return;
        // Only allow alphanumeric input
        const newOtp = [...otp];
        newOtp[index] = value;
        setOtp(newOtp);

        // Focus the next input field
        if (value !== '' && index < 5) {
            inputsRef.current[index + 1].focus();
        }
    };

    // Render OTP input boxes
    const renderOtpInputs = () => {
        return otp.map((data, index) => (
            <input
                key={index}
                type="text"
                maxLength="1"
                value={otp[index]}
                // onChange={(e) => handleOtpChange(index, e.target.value)}
                onChange={(e) => handleOtpChange(index, e.target.value.replace(/[^0-9]/g, ''))} // Allow only numbers
                onKeyDown={(e) => handleKeyDown(index, e)}
                ref={(el) => (inputsRef.current[index] = el)}
                className="otp-input"
                inputMode="numeric" // Suggest numeric keyboard for mobile devices
                pattern="[0-9]*" // Ensure that only numeric characters are accepted
            />
        ));
    };

    const showResendOtp = () => {
        setShowErrorPopup(false);
        setTimerExpired(true);
    }

    const togglePasscodeVisibility = (isVisible) => {
        setIsPasscodeVisible(isVisible);
    };

    // Function to toggle visibility
    const toggleConfirmPasscodeVisibility = (isVisible) => {
        setIsConfirmPasscodeVisible(isVisible);
    };

    // Handle key down (Enter or Space) to show passcode
    const handleKeyDown = (event,index) => {
        if (event.key === 'Enter' || event.key === ' ') {
            toggleConfirmPasscodeVisibility(true); // Show passcode
            togglePasscodeVisibility(true);
        }
        if (event.key === 'Backspace' && otp[index] === '' && index > 0) {
            inputsRef.current[index - 1].focus();
        }
        if (event.key === 'ArrowLeft' && index > 0) {
            inputsRef.current[index - 1].focus();
        }
        if (event.key === 'ArrowRight' && index < 5) {
            inputsRef.current[index + 1].focus();
        }
    };

    // Handle key up to hide passcode
    const handleKeyUp = (event) => {
        if (event.key === 'Enter' || event.key === ' ') {
            toggleConfirmPasscodeVisibility(false); // Hide passcode
            togglePasscodeVisibility(false);
        }
    };

    // Function to handle passcode input change
    const handlePasscodeChange = (event) => {
        const passcodeValue = event.target.value;
        // setPasscode(passcodeValue);
        const alphanumericRegex = /^[a-zA-Z0-9]*$/;
        const regex = /^(?=.*[A-Za-z])(?=.*\d)[A-Za-z\d]+$/;

        if (alphanumericRegex.test(passcodeValue)) {
            setPasscode(passcodeValue);
        }

        // Check if the confirm passcode matches the new passcode value
        if (confimPasscode !== passcodeValue && confimPasscode !== "") {
            setError("Passcodes do not match");
        }
        else if (!regex.test(passcodeValue) && passcodeValue.length > 5) {
            setError("Passcode should contain both Alphabets and Numbers.");
        } else {
            setError("");
        }
    };

    // Function to handle confirm passcode input change
    const handleConfirmPasscodeChange = (event) => {
        const confirmPasscodeValue = event.target.value;
        // setConfimPasscode(confirmPasscodeValue);
        const alphanumericRegex = /^[a-zA-Z0-9]*$/;

        if (alphanumericRegex.test(confirmPasscodeValue)) {
            setConfimPasscode(confirmPasscodeValue);
        }
        // Check if the confirm passcode matches the passcode value
        const regex = /^(?=.*[A-Za-z])(?=.*\d)[A-Za-z\d]+$/;

        if (passcode !== confirmPasscodeValue && passcode !== "") {
            setError("Entered Passcodes do not match");
        }
        else if (!regex.test(confirmPasscodeValue) && confirmPasscodeValue.length > 5) {
            setError("Passcode should contain both Alphabets and Numbers.");
        } else {
            setError("");
        }
    };

    return (
        <div>
            <Navbar header="IB Smart Office - eNote" />
            <Sidebar />
            {/* Added top navigation for menu --> 23-09 */}
            <MenuNavContainer isMenuPage />
            <div className="container largeScreen text-center">
                <div className="SectionHeads row mobileSectionHeads">Create/Reset Passcode</div>
                {!passcodeSent && !passcodeVerified && (
                    <div className='card addPasscode'>
                        <form onSubmit={handleAddPasscode} className="form-container">
                            <div className="form-group">
                                <div className="passcodeHint">Your Passcode must consist of 6 alpha-numeric characters</div>
                                <label>Enter a Passcode:</label>

                                <div className="passcode-input-container">
                                    <input
                                        type={isPasscodeVisible ? "text" : "password"}
                                        id="passcode"
                                        value={passcode}
                                        onChange={handlePasscodeChange}
                                        className="passcode-input"
                                        maxLength="6"
                                        autoComplete="off"
                                        pattern="\d*"
                                        title="Please enter 6-characters, Combination of Alphabets and Numbers"
                                    />
                                    <span className="eye-icon"
                                        tabIndex="0" // Makes the span focusable for keyboard events
                                        onMouseDown={() => togglePasscodeVisibility(true)}    // Show on mouse down
                                        onMouseUp={() => togglePasscodeVisibility(false)}     // Hide on mouse up
                                        onKeyDown={(e) => handleKeyDown(e)}                   // Show on key down (Enter or Space)
                                        onKeyUp={(e) => handleKeyUp(e)}   
                                        onTouchStart={() => togglePasscodeVisibility(true)}   // Show on touch start (mobile)
                                        onTouchEnd={() => togglePasscodeVisibility(false)} >
                                        <img
                                            src={isPasscodeVisible ? view : hide}
                                            alt={isPasscodeVisible ? "Hide passcode" : "Show passcode"}
                                        />
                                    </span>
                                </div>
                            </div>
                            <div className="form-group">
                                <label htmlFor="confirm-passcode">Confirm Passcode:</label>
                                <div className="passcode-input-container">
                                    <input
                                        type={isConfirmPasscodeVisible ? "text" : "password"}
                                        id="confirm-passcode"
                                        value={confimPasscode}
                                        onChange={handleConfirmPasscodeChange}
                                        className="passcode-input"
                                        maxLength="6"
                                        pattern="\d*"
                                        autoComplete="off"
                                        title="Please enter 6-characters, Combination of Alphabets and Numbers"
                                    />
                                    <span
                                        className="eye-icon-view-confirm"
                                        tabIndex="0" // Makes the span focusable for keyboard events
                                        onMouseDown={() => toggleConfirmPasscodeVisibility(true)}    // Show on mouse down
                                        onMouseUp={() => toggleConfirmPasscodeVisibility(false)}     // Hide on mouse up
                                        onKeyDown={(e) => handleKeyDown(e)}                          // Show on key down (Enter or Space)
                                        onKeyUp={(e) => handleKeyUp(e)}  
                                        onTouchStart={() => toggleConfirmPasscodeVisibility(true)}   // Show on touch start (mobile)
                                        onTouchEnd={() => toggleConfirmPasscodeVisibility(false)}    // Hide on key up
                                    >
                                        <img
                                            src={isConfirmPasscodeVisible ? view : hide}
                                            alt={isConfirmPasscodeVisible ? "Hide confirm passcode" : "Show confirm passcode"}
                                        />
                                    </span>
                                </div>

                                {error && <p className="confirmerror">{error}</p>}
                            </div>
                            <Button
                                type="submit"
                                className="FormButtons k-button k-button-md k-rounded-md k-button-solid k-button-solid-base"
                                onClick={handleAddPasscode}
                            >
                                <span className="k-icon k-font-icon k-i-launch cursor allIconsforPrimary-btn"></span>
                                Submit
                            </Button>
                        </form>
                    </div>
                )}
                {passcodeSent && !passcodeVerified && (
                    <div className='card mobileCSSStyle'>
                        {/* Changed the OTP content -- feedback --> Kavya (25-07) */}
                        <p>OTP has been sent to your mobile number & email address.</p>
                        <p>Please enter your OTP</p>
                        <div className="otp-container">{renderOtpInputs()}</div>
                        {!timerExpired && (
                            <>
                                {!saveButtonVisible && (
                                    <Button onClick={handleVerifyPasscode} className="passcodebutton FormButtons k-button k-button-md k-rounded-md k-button-solid k-button-solid-base">
                                        Verify OTP
                                    </Button>
                                )}
                                {verifyContentHide && <p className='timer'>Time remaining: <b>{timer}</b> seconds</p>}
                            </>
                        )}
                        {timerExpired && (
                            <div>
                                <p className='resendPara'>OTP has been expired.</p>
                                <Button onClick={handleResendPasscode} className="passcodebutton FormButtons k-button k-button-md k-rounded-md k-button-solid k-button-solid-base">Resend OTP</Button>
                            </div>
                        )}
                    </div>
                )}
                {passcodeVerified && (
                    <div className='card mobileCSSStyle'>
                        <p className='otpverified'>OTP verified successfully!</p>
                        {saveButtonVisible && (
                            <Button onClick={handleSavePasscode} className="passcodebutton FormButtons k-button k-button-md k-rounded-md k-button-solid k-button-solid-base">
                                Save Passcode
                            </Button>
                        )}
                    </div>
                )}

                {showSuccessPopup && (
                    <Dialog
                        title="Success Popup"
                        onClose={() => showSuccessPopup(false)}
                        className="dialogcontent_Refer"
                    >
                        <p className='dialogcontent_'>
                            Passcode saved successfully!
                        </p>
                        <DialogActionsBar>
                            <Button
                                onClick={handleCloseSuccessPopup}
                                className="notifyDailogOkBtn k-button k-button-md k-rounded-md k-button-solid k-button-solid-base"
                            >
                                <span className="k-icon k-font-icon k-i-redo cursor allIconsforPrimary-btn"></span>
                                Ok
                            </Button>
                        </DialogActionsBar>
                    </Dialog>
                )}
                {showErrorPopup && (
                    <Dialog
                        title="Alert!"
                        onClose={() => setShowErrorPopup(false)}
                        className="dialogcontent_Refer"
                    >
                        <p className='dialogcontent_'>
                            Verification failed. Please try again
                        </p>
                        <DialogActionsBar>
                            <Button
                                onClick={showResendOtp}
                                className="notifyDailogOkBtn k-button k-button-md k-rounded-md k-button-solid k-button-solid-base"
                            >
                                <span className="k-icon k-font-icon k-i-redo cursor allIconsforPrimary-btn"></span>
                                Ok
                            </Button>
                        </DialogActionsBar>
                    </Dialog>
                )}
            </div>
            {/* {isLoading && <PageLoader />} */}
            <div className="pgFooterContainer">
                <Footer />
            </div>
        </div>
    );
};