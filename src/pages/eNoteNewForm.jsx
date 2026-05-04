import React, { useState, useEffect } from "react";
// import kendo components
import { Form, FormElement, FieldWrapper, Field } from "@progress/kendo-react-form";
import { Input, TextArea, Switch } from "@progress/kendo-react-inputs";
import { DropDownList, MultiColumnComboBox } from "@progress/kendo-react-dropdowns";
import { Button } from "@progress/kendo-react-buttons";
import { Label, Hint } from "@progress/kendo-react-labels";
import { Dialog, DialogActionsBar } from "@progress/kendo-react-dialogs";
// import kendo icons
import { SvgIcon } from "@progress/kendo-react-common";
import { infoCircleIcon } from "@progress/kendo-svg-icons";
import { filePdfIcon, fileWordIcon, fileImageIcon, fileTxtIcon, fileDataIcon, fileIcon, xIcon } from "@progress/kendo-svg-icons";
// import external components
import { Link, useParams, useNavigate } from "react-router-dom";
import { useMsal, useAccount } from "@azure/msal-react";
import DateObject from "react-date-object";
import Clock from 'react-live-clock';
// import internal components
import Navbar from "../components/navbar";
import Footer from "../components/footer";
import { Sidebar } from "../components/sidebar";
import { MenuNavContainer } from "../components/menu";
import { TableDraggableRows } from "../components/tableDragRows";
import { PageLoader } from "../components/pageLoader";
import { API_BASE_URL, API_ENDPOINTS, loginRequest, API_COMMON_HEADERS, charValidation } from "../config";
import { getAccessToken, useTabContext } from "../App";
import Unauthorized from "../components/Unauthorized";
import useAutosave from "../hooks/useAutoSave.hook";
// Mobile responsive css and css styles
import "../styles/responsiveDesign.css";
import "bootstrap/dist/css/bootstrap.min.css";
import "../styles/forms.css";

const CustomDialogTitleBar = () => {
  return (
    <div className="custom-title cstDailogIbHeader">
      <SvgIcon icon={infoCircleIcon} /> Alert!
    </div>
  );
};

const CustomConfirmDialogTitleBar = () => {
  return (
    <div className="custom-title cstDailogIbHeader">
      <span className="k-icon k-font-icon k-i-borders-show-hide cursor allIconsforPrimary-btn"></span> Confirmation
    </div>
  );
};

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
    header: "Employee Id ",
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
    header: "Employee Id",
    width: "70px",
  },
];
export const ReorderContext = React.createContext({
  reorder: () => { },
  dragStart: () => { },
});

export const ENoteNewForm = () => {
  const valueRender = (element, value, fieldName) => {
    const clearValue = (e) => {
      e.stopPropagation();
      e.preventDefault();
      setState(prevState => ({
        ...prevState,
        [fieldName]: '', // Clear the value of the specified field
      }));
    };
    if (!value) {
      return element;
    }
    const children = [
      <span key={1} className="_valueRender">
        {element.props.children}
      </span>,
      <SvgIcon icon={xIcon} onClick={clearValue} />
    ];
    return React.cloneElement(element, {
      ...element.props
    }, children);
  };

  const { setTab } = useTabContext();
  const navigate = useNavigate();
  const [noteto, setNoteTo] = useState([]);
  const [natureofnote, setNatureOfNote] = useState([]);
  const [notetype, setNoteTypeData] = useState([]);
  const [fintype, setFinType] = useState([]);
  const [natureofapprsanc, setNatureofApprSanc] = useState([]);
  const [noteapproverData, setNoteApproverData] = useState([]);
  const [notereviewerData, setNoteReviewerData] = useState([]);
  const [noteComments, setNoteComments] = useState([]);
  const [showNatureOfApprSanc, setShowNatureOfApprSanc] = useState(false);
  const [showTypeOfFinNote, setshowTypeOfFinNote] = useState(false);
  const [showAmount, setshowAmount] = useState(false);
  const [showPurpose, setshowPurpose] = useState(false);
  const [showPurpose1, setshowPurpose1] = useState(false);
  const [showPurpose2, setshowPurpose2] = useState(false);
  const [departmentError] = useState("");
  const [noteBorderColor, setNoteBorderColor] = useState("rgba(0, 0, 0, 0.12)");
  const [isLoading, setIsLoading] = useState(false);
  const [successMessage, setSuccessMessage] = useState("");
  const urlParam = useParams();
  const [noteId] = useState(urlParam.id);
  const [draftId, setDraftId] = useState(noteId === "new" ? 0 : noteId);
  const [redirectTo] = useState("/datagridpage");
  const [allOrgUsers, setAllOrgUsers] = useState([]);
  const [searchBorderColor, setSearchBorderColor] = useState("rgba(0, 0, 0, 0.12)");
  const [subjectBorderColor, setSubjectBorderColor] = useState("rgba(0, 0, 0, 0.12)");
  const [natureofapprsancBorderColor, setNatureofapprsancBorderColor] = useState("rgba(0, 0, 0, 0.12)");
  const [notetypeBorderColor, setNotetypeBorderColor] = useState("rgba(0, 0, 0, 0.12)");
  const [typeoffinBorderColor, setTypeOfFinBorderColor] = useState("rgba(0, 0, 0, 0.12)");
  const [amountBorderColor, setAmountBorderColor] = useState("rgba(0, 0, 0, 0.12)");
  const [natureofnoteBorderColor, setNatureOfNoteBorderColor] = useState("rgba(0, 0, 0, 0.12)");
  const [purposeBorderColor, setPurposeBorderColor] = useState("rgba(0, 0, 0, 0.12)");
  const [purposeOthersBorderColor, setPurposeOthersBorderColor] = useState("rgba(0, 0, 0, 0.12)");
  const [purpose1BorderColor, setPurpose1BorderColor] = useState("");
  const [purpose2BorderColor, setPurpose2BorderColor] = useState("");
  const [selectedReveiwer, setSelectedReviewer] = useState(null);
  const [combovalueApprover, setComboValueApprover] = useState(null);
  const [visible, setVisible] = useState(false);
  const [visibleCancelCfrm, setVisibleCancelCfrm] = useState(false);
  const [alertvisible, setAlertVisible] = useState(false);
  const [visiblesave, setVisibleSave] = useState(false);
  const [validationErrors, setValidationErrors] = useState(false);
  const [errorMessages, setErrorMessages] = useState([]);
  const [isSecretaryExist, setisSecretaryExist] = useState(false);
  const [showNotification, setShowNotification] = useState(false);
  const [notificationMsg, setNotificationMsg] = useState(false);
  const [enumsObj, setEnumsObj] = useState(null);
  const [isNewForm] = useState(noteId === "new");
  const [amount, setAmount] = useState("");
  const [purpose, setPurpose] = useState("");
  const [purposeOthers, setPurposeOthers] = useState("");
  const [drpOptPurposeApproval, setDrpOptPurposeApproval] = useState([]);
  const [drpOptPurposeInformation, setDrpOptPurposeInformation] = useState([]);
  const [supportingDocError, setSupportingError] = useState("");
  const [notepath, setNotePath] = useState(null);
  const [wordPath, setWordPath] = useState(null);
  const [isConfidential, setIsConfidential] = useState(false);
  const [showConfidential, setShowConfidential] = useState(false); // New state for showing Confidential Note toggle
  const max = 250;
  const isMobile = window.innerWidth <= 768;
  const [state, setState] = useState({
    Department: "",
    NoteTo: "",
    Subject: "",
    NatureOfNote: "",
    NatureOfApprSanc: "",
    NoteType: "",
    TypeOfFinNote: "",
    Amount: "",
    Search: "",
    Purpose: "",
    PurposeApproval: "",
    PurposeInformation: "",
    notePdfPath: "",
    noteWordPath: "",
    supportingdocument: "",
    createdBy: "",
    status: "",
    statusMsg: "",
    createdDate: "",
    isConfidential: false
  });
  const [filesInfo, SetFileinfo] = useState([]);
  const [wordandPdfWarring, SetWordPDFInfowarring] = useState({
    PDFInfo: {
      fileExtension: "",
      fileName: "",
      warningMsg: "",
      filePath: "",
      isValid: false,
    },
    wordInfo: {
      fileExtension: "",
      fileName: "",
      warningMsg: "",
      filePath: "",
      isValid: false,
    },
  });

  const [supportDocWarning, SetSupportDocWarning] = useState({});
  const { instance, accounts } = useMsal();
  const account = useAccount(accounts[0] || {});
  const [userDepartment, setUserDepertment] = useState("");
  const [userDesignation, setUserDesignation] = useState("");
  const [isSaveAndSubmitActionCompleted, setisSaveAndSubmitActionCompleted] = useState(false);
  const [departmentList, setDepartmentList] = useState([]);
  const [isVieweEligible, setIsVieweEligible] = useState("true");
  const [purposeErrorMessage, setPurposeErrorMessage] = useState("");
  const [subjectErrorMessage, setSubjectErrorMessage] = useState("");
  const [searchErrorMessage, setSearchErrorMessage] = useState("");
  const [fileKey, setFileKey] = useState(Date.now()); // Unique key for file input
  const [orgEmployees, setOrgEmployees] = useState([]);
  const maxPurpose = 500;

  useEffect(() => {
    if (!isNewForm) {
      isUserViewEligible();
    }
    setIsLoading(true);
    getUserDepartment();
  }, []);

  const isUserViewEligible = async () => {
    let isUserViewEligible = true;
    try {
      const accessToken = await getAccessToken({ ...loginRequest, account }, instance);
      const response = await fetch(`${API_BASE_URL}${API_ENDPOINTS.GET_NoteViewEligibility}?userPrincipalName=${accounts[0].username}&NoteId=${draftId}`, {
        method: "POST",
        headers: {
          ...API_COMMON_HEADERS,
          Authorization: `Bearer ${accessToken}`,
        },
      });

      const result = await response.text();
      if (result) {
        isUserViewEligible = result;

        if (result === "false") {
          setIsLoading(false)
        }
        setIsVieweEligible(result)
      }
      return isUserViewEligible;
    } catch (error) {
      setIsLoading(false)
      console.error("Error fetching data:", error);
    }
  };

  //Get login users department and designation
  const getUserDepartment = async () => {
    try {
      const accessToken = await getAccessToken({ ...loginRequest, account }, instance);

      const obj = await fetch(
        `${API_BASE_URL}${API_ENDPOINTS.GET_UserDetailsByPrincipalName(accounts[0].username)}`, {
        method: "GET",
        headers: { ...API_COMMON_HEADERS, Authorization: `Bearer ${accessToken}` }
      }
      );
      const departmentDetails = await obj.json();
      // Commentted UserDepartment State in onload
      setUserDepertment(departmentDetails[0].department);
      setUserDesignation(departmentDetails[0].jobTitle);

      await getDefaultConfigData(departmentDetails[0].department);
    } catch (error) {
      console.error("Error fetching user details:", error);
    }
  };

  // Change 13/05 API call for department list
  const getDepartmentList = async () => {
    const accessToken = await getAccessToken({ ...loginRequest, account }, instance);

    const obj = await fetch(`${API_BASE_URL}${API_ENDPOINTS.GET_Departments_List}`
      , {
        method: "GET",
        headers: { ...API_COMMON_HEADERS, Authorization: `Bearer ${accessToken}` }
      }
    );
    const departmentDetailsList = await obj.json();
    setDepartmentList(departmentDetailsList)
  }

  //Find secretary exists for the added Approvers/Reviewers
  const findSecrateries = async (apprObj, userDepartment) => {
    let param = apprObj.map((x) => ({
      approverType: x.approverType,
      approverEmail: x.approverEmail,
    }));
    const accessToken = await getAccessToken({ ...loginRequest, account }, instance);
    let secretaryExists = false;
    if (userDepartment) {
      try {
        await fetch(
          `${API_BASE_URL}${API_ENDPOINTS.eNote_GetNoteSecretaryofApprover}`,
          {
            method: "POST",
            headers: { ...API_COMMON_HEADERS, Authorization: `Bearer ${accessToken}` },
            body: JSON.stringify({ noteApproversMasterlst: param, departmentName: userDepartment }),
          }
        )
          .then((response) => {
            return response.json();
          })
          .then((data) => {
            setisSecretaryExist(data.secretaryExist);
            secretaryExists = data.secretaryExist;
          });
      } catch (error) {
        console.error("Error fetching data from API:", error);
      }
    }
    return secretaryExists;
  };

  //Get all the dropdowns and pre configured data
  // const getDefaultConfigData = async (userDepartment) => {
  //   try {
  //     const accessToken = await getAccessToken({ ...loginRequest, account }, instance);
  //     //get enum objects
  //     const dropdowns = await fetch(`${API_BASE_URL}${API_ENDPOINTS.GET_DROPDOWNDATA}`, {
  //       method: "GET",
  //       headers: { ...API_COMMON_HEADERS, Authorization: `Bearer ${accessToken}` }
  //     });
  //     const dropdownslist = await dropdowns.json();
  //     getDepartmentList();
  //     setEnumsObj(dropdownslist);
  //     setNoteTo(dropdownslist.noteTo);
  //     setNatureOfNote(dropdownslist.natureofNote);
  //     setNoteTypeData(dropdownslist.noteType);
  //     setFinType(dropdownslist.financialType);
  //     setNatureofApprSanc(dropdownslist.natureOfApprovalOrSanction);
  //     setDrpOptPurposeApproval(dropdownslist.PurposeonApproval);
  //     setDrpOptPurposeInformation(dropdownslist.PurposeonInformation);

  //     if (!(isNewForm)) {
  //       await fetcheNotedetails(userDepartment, dropdownslist.PurposeonApproval);
  //     }
  //     setIsLoading(false);

  //     if (isNewForm) {
  //       const approverresponse = await fetch(
  //         `${API_BASE_URL}${API_ENDPOINTS.eNote_GetNoteApprovers(userDepartment)}`, {
  //         method: "GET",
  //         headers: { ...API_COMMON_HEADERS, Authorization: `Bearer ${accessToken}` }
  //       }
  //       );

  //       const reviewerresponse = await fetch(
  //         `${API_BASE_URL}${API_ENDPOINTS.eNote_GetNoteReviewer(userDepartment)}`, {
  //         method: "GET",
  //         headers: { ...API_COMMON_HEADERS, Authorization: `Bearer ${accessToken}` }
  //       });

  //       const noteapprover = await approverresponse.json();
  //       const notereviewer = await reviewerresponse.json();
  //       if (noteapprover) {
  //         setNoteApproverData(noteapprover);
  //         await findSecrateries(noteapprover, userDepartment);
  //       }
  //       if (notereviewer)
  //         setNoteReviewerData(notereviewer);
  //       setIsLoading(false);
  //     }

  //   } catch (error) {
  //     console.error("Error fetching data:", error);
  //     setIsLoading(false);
  //   }
  // };

  const getDefaultConfigData = async (userDepartment) => {
    try {
      const accessToken = await getAccessToken({ ...loginRequest, account }, instance);

      // get enum objects (dropdown data)
      const dropdowns = await fetch(
        `${API_BASE_URL}${API_ENDPOINTS.GET_DROPDOWNDATA}`,
        {
          method: "GET",
          headers: { ...API_COMMON_HEADERS, Authorization: `Bearer ${accessToken}` }
        }
      );
      const dropdownslist = await dropdowns.json();

      getDepartmentList();
      setEnumsObj(dropdownslist);
      setNoteTo(dropdownslist.noteTo);
      setNatureOfNote(dropdownslist.natureofNote);
      setNoteTypeData(dropdownslist.noteType);
      setFinType(dropdownslist.financialType);
      setNatureofApprSanc(dropdownslist.natureOfApprovalOrSanction);
      setDrpOptPurposeApproval(dropdownslist.PurposeonApproval);
      setDrpOptPurposeInformation(dropdownslist.PurposeonInformation);

      // 🔹 check if user department is in ConfidentialDepartments
      const confidentialDepts = dropdownslist?.ConfidentialDepartments || [];

      const userDeptRes = await fetch(`${API_BASE_URL}${API_ENDPOINTS.Find_Super_Dept_Admin(accounts[0]?.username)}`, {
        method: "GET",
        headers: { ...API_COMMON_HEADERS, Authorization: `Bearer ${accessToken}` }
      });
      const userDeptObj = await userDeptRes.json();

      if (userDeptObj?.departmentName && confidentialDepts.includes(userDeptObj.departmentName)) {
        setShowConfidential(true); // 👈 add a new state for showing the Confidential Note toggle
      } else {
        setShowConfidential(false);
      }

      if (!isNewForm) {
        await fetcheNotedetails(userDepartment, dropdownslist.PurposeonApproval);
      }
      setIsLoading(false);

      if (isNewForm) {
        const approverresponse = await fetch(
          `${API_BASE_URL}${API_ENDPOINTS.eNote_GetNoteApprovers(userDepartment)}`,
          {
            method: "GET",
            headers: { ...API_COMMON_HEADERS, Authorization: `Bearer ${accessToken}` }
          }
        );

        const reviewerresponse = await fetch(
          `${API_BASE_URL}${API_ENDPOINTS.eNote_GetNoteReviewer(userDepartment)}`,
          {
            method: "GET",
            headers: { ...API_COMMON_HEADERS, Authorization: `Bearer ${accessToken}` }
          }
        );

        const noteapprover = await approverresponse.json();
        const notereviewer = await reviewerresponse.json();

        if (noteapprover) {
          setNoteApproverData(noteapprover);
          await findSecrateries(noteapprover, userDepartment);
        }
        if (notereviewer) setNoteReviewerData(notereviewer);

        setIsLoading(false);
      }
    } catch (error) {
      console.error("Error fetching data:", error);
      setIsLoading(false);
    }
  };


  //Get eNote data for Edit form on load
  const fetcheNotedetails = async (userDepartment, purposeApprovalDrpOpt) => {
    setIsLoading(true);
    try {
      const accessToken = await getAccessToken({ ...loginRequest, account }, instance);

      const response = await fetch(
        `${API_BASE_URL}${API_ENDPOINTS.eNote_GetGeneralDetails}`,
        {
          method: "POST",
          headers: { ...API_COMMON_HEADERS, Authorization: `Bearer ${accessToken}` },
          body: JSON.stringify({ noteId: noteId }),
        }
      );
      if (!response.ok) {
        // Handle error, for example set an error state
        console.error("Error fetching data:", response.statusText);
        return;
      }
      const data = await response.json();

      setState({
        ...data,
        status: data.status,
        NoteTo: data.strNoteTo,
        Amount: data.amount || "",
        NatureOfNote: data.strNatureofNote === "0" ? "" : data.strNatureofNote,
        Subject: data.subject || "",
        NatureOfApprSanc: data.strNatureOfApprovalOrSanction === "0" ? "" : data.strNatureOfApprovalOrSanction,
        NoteType: data.strNoteType === "0" ? "" : data.strNoteType,
        FinancialType: data.strFinancialType === "0" ? "" : data.strFinancialType,
        TypeOfFinNote: data.strFinancialType === "0" ? "" : data.strFinancialType,
        Search: data.searchKeyword || "",
        Purpose: data.purpose || "",
        PurposeApproval: data.purpose === "0" ? "" : data.purpose,
        PurposeInformation: data.purpose === "0" ? "" : data.purpose,
        NotePdfFileName: data.purpose === "0" ? "" : data.purpose,
        NotePdfPath: data.notePdfPath || "",
        supportingdocument: data.noteSupportingDocumentsDTO || "",
        noteApproversDTO: data.noteApproversDTO || "",
        createdBy: data.createdBy,
        statusMsg: data.strNoteStatus,
        createdDate: new DateObject(new Date(data.createdDate)).format("DD-MMM-YYYY hh:mm A"), //seconds hand was removed 
        isConfidential: data.isConfidential ?? false
      });
      setIsConfidential(data.isConfidential ?? false);

      setUserDepertment(data?.departmentName ? data?.departmentName : "");
      setAmount(data.amount);

      if (data.strNatureofNote === "Approval" && purposeApprovalDrpOpt.find(x => (x.dValue === data.purpose)) === undefined) {
        setPurpose("Others");
        setPurposeOthers(data.purpose);
      } else {
        setPurpose(data.purpose);
      }
      setNoteComments(data.noteApproverCommentsDTO);
      setShowNatureOfApprSanc(data.strNatureofNote === "Sanction" || data.strNatureofNote === "Approval");
      setshowPurpose(data.strNatureofNote === "Sanction" || data.strNatureofNote === "Ratification");
      setshowPurpose1(data.strNatureofNote === "Approval");
      setshowPurpose2(data.strNatureofNote === "Information");
      setshowTypeOfFinNote(data.strNoteType === "Financial");
      setshowAmount(data.strNoteType === "Financial");

      SetWordPDFInfowarring({
        PDFInfo: {
          fileName: data?.notePdfFileName,
          warningMsg: "",
          filePath: data?.notePdfPath,
          isValid: (data?.notePdfFileName === "" || data?.notePdfFileName === null) ? false : true
        },
        wordInfo: {
          fileName: data?.noteWordFileName,
          warningMsg: "",
          filePath: data?.noteWordPath,
          isValid: (data?.noteWordFileName === "" || data?.noteWordFileName === null) ? false : true
        },
      });
      setNotePath(data?.notePdfPath);
      setWordPath(data?.noteWordPath);
      SetFileinfo(data?.noteSupportingDocumentsDTO);
      const fetchNoteApprover = data.noteApproversDTO?.filter((obj) => obj.approverType === 2 && obj.isActive);
      if (fetchNoteApprover?.length > 0) {
        await findSecrateries(fetchNoteApprover, data?.departmentName ? data?.departmentName : "");
      }
      setNoteApproverData(fetchNoteApprover);

      setNoteReviewerData(
        data.noteApproversDTO?.filter((obj) => obj.approverType === 1 && obj.isActive)
      );
    } catch (error) {
      console.error("Error fetching data:", error.message);
    }
    setIsLoading(false);
  };

  // Delete Approvers from Approvers/Reviewers table
  const deletereviewer = (data) => {
    if (!data) return;

    if (data.approverType === 1) {
      let updatedData;
      if (data.noteApproverMasterId) {
        // Existing reviewer: delete by noteApproverMasterId
        updatedData = notereviewerData.filter(
          (row) => row.noteApproverMasterId !== data.noteApproverMasterId
        );
      } else {
        // New reviewer: delete by email + order (best effort)
        updatedData = notereviewerData.filter(
          (row) =>
            !(
              row.approverEmail === data.approverEmail &&
              row.approverOrder === data.approverOrder &&
              !row.noteApproverMasterId // Only target new rows
            )
        );
      }
      // Reorder
      updatedData.forEach((item, index) => {
        item.approverOrder = index + 1;
      });
      setNoteReviewerData(updatedData);
    }

    if (data.approverType === 2) {
      let updatedData;
      if (data.noteApproverMasterId) {
        updatedData = noteapproverData.filter(
          (row) => row.noteApproverMasterId !== data.noteApproverMasterId
        );
      } else {
        updatedData = noteapproverData.filter(
          (row) =>
            !(
              row.approverEmail === data.approverEmail &&
              row.approverOrder === data.approverOrder &&
              !row.noteApproverMasterId
            )
        );
      }
      updatedData.forEach((item, index) => {
        item.approverOrder = index + 1;
      });
      setNoteApproverData(updatedData);
    }
  };

  //  Approvers orderChange from Approvers/Reviewers table   
  const orderChange = (data) => {
    if (data.some(obj => obj.approverType === 1)) {
      setNoteReviewerData(data)
    }
    if (data.some(obj => obj.approverType === 2)) {
      setNoteApproverData(data);
    }
  }

  // Autosave functionaites  is implemented here
  useAutosave(() => {
    if (!isSaveAndSubmitActionCompleted && isNewForm) {
      handleAutoSave();
    }
  }, 3 * 60 * 1000);

  /* -------------------Form fields handlers START ---------------------- */
  /* Data Handles */
  const handleNote = (event) => {
    setState((prevState) => ({ ...prevState, NoteTo: event.target.value }));
    setNoteBorderColor("");
  };

  const validateInput = (input) => {
    const allowedChars = charValidation;
    return allowedChars.test(input);
  };
  //  handle subject change


  // handle  nature of note change
  const handleNatureOfNoteChange = (event) => {
    setState((prevState) => ({ ...prevState, NatureOfNote: event.target.value }));
    setNatureOfNoteBorderColor("");
    const selectedValue = event.target.value;
    if (selectedValue === "Sanction" || selectedValue === "Ratification")
      setState((prevState) => ({ ...prevState, PurposeApproval: "", PurposeInformation: "", NatureOfApprSanc: "" }));

    if (selectedValue === "Approval")
      setState((prevState) => ({ ...prevState, Purpose: "", PurposeInformation: "", NatureOfApprSanc: "" }));

    if (selectedValue === "Information")
      setState((prevState) => ({ ...prevState, PurposeApproval: "", Purpose: "", NatureOfApprSanc: "" }));

    setShowNatureOfApprSanc(selectedValue === "Sanction" || selectedValue === "Approval");
    setshowPurpose(selectedValue === "Sanction" || selectedValue === "Ratification");
    setshowPurpose1(selectedValue === "Approval");
    setshowPurpose2(selectedValue === "Information");
    setPurpose('');
    setPurposeOthers('');
  };

  // handle  Natureb Of ApprSanc Change
  const handleNatureOfApprSancChange = (event) => {
    setState((prevState) => ({ ...prevState, NatureOfApprSanc: event.target.value }));
    setNatureofapprsancBorderColor("");
  };

  // handle NoteType change
  const handleNoteType = (event) => {
    setState((prevState) => ({ ...prevState, NoteType: event.target.value }));
    if (!(event.target.value === "Financial"))
      setState((prevState) => ({ ...prevState, TypeOfFinNote: "", Amount: 0 }));
    setNotetypeBorderColor("");
    setshowTypeOfFinNote(event.target.value === "Financial");
    setshowAmount(event.target.value === "Financial");
  };

  // handle  Type Of Financial
  const handleTypeOfFin = (event) => {
    setState((prevState) => ({ ...prevState, TypeOfFinNote: event.target.value }));
    setTypeOfFinBorderColor("");
  };

  // handle amoount change
  const handleAmount = (event) => {

    if (/^\d*\.?\d{0,2}$/.test(event.target.value)) {
      setState((prevState) => ({ ...prevState, Amount: event.target.value }));
      setAmount(event.target.value);
      setAmountBorderColor("");
    }
  };

  const handleSubjectChange = (event) => {
    const value = event.target.value;

    setState((prevState) => ({ ...prevState, Subject: value }));
    setSubjectBorderColor('');
    setSubjectErrorMessage('');
  };
  // handle search change
  const handleSearchChange = (event) => {
    const value = event.target.value;

    setState((prevState) => ({ ...prevState, Search: value }));
    setSearchBorderColor(''); // Reset border color if valid
    setSearchErrorMessage(''); // Clear error message if valid
  };

  //handle  purpose change
  const handlePurpose = (e) => {
    const value = e.target.value;
    setPurpose(value);
    setPurposeBorderColor('');
    setPurposeErrorMessage('');
  };

  // handle Purpose Info Others change
  const handlePurposeInfoOthers = (event) => {
    setPurposeOthers(event.target.value);
    setPurposeOthersBorderColor("");
  };

  //  handle Purpose Approval change
  const handlePurposeApproval = (event) => {
    setState((prevState) => ({ ...prevState, PurposeApproval: event.target.value }));
    setPurpose(event.target.value);
    setPurposeOthers('');
    setPurpose1BorderColor("");
  };

  // handle Purpose Information change
  const handlePurposeInformation = (event) => {
    setState((prevState) => ({ ...prevState, PurposeInformation: event.target.value }));
    setPurpose(event.target.value);
    setPurpose2BorderColor("");
  };
  /* -------------------Form fields handlers END ---------------------- */

  /* Form mandatory fields and file validation handler */
  const validateForm = (isApproverValid, isReviewerValid) => {
    const errors = {};
    let totalFileSize = 0;

    const inValidFileNames = Object.keys(supportDocWarning).filter(
      (fileName) => supportDocWarning[fileName]?.isValid === false
    );

    if (filesInfo.length > 0) {
      const validMultipleFile = filesInfo.filter(obj => {
        if (!inValidFileNames.includes(obj.supportingDocumentFileName)) {
          return obj;
        }
      });
      validMultipleFile.forEach(obj => {
        totalFileSize += obj.supportingDocumentPathLength;
      });
    }

    // === Standard field validations ===
    if (!state.NoteTo) {
      errors["NoteTo"] = "NoteTo";
      setNoteBorderColor("#f31700");
    } else {
      setNoteBorderColor("");
    }

    if (!state.Subject.trim()) {
      errors["Subject"] = "Subject";
      setSubjectBorderColor("#f31700");
    } else {
      setSubjectBorderColor("");
    }

    if (!state.NatureOfNote) {
      errors["NatureOfNote"] = "Nature Of Note";
      setNatureOfNoteBorderColor("#f31700");
    } else {
      setNatureOfNoteBorderColor("");
    }

    if (showNatureOfApprSanc && !state.NatureOfApprSanc) {
      errors["NatureOfApprSanc"] = "Nature Of Appr/Sanc";
      setNatureofapprsancBorderColor("#f31700");
    } else {
      setNatureofapprsancBorderColor("");
    }

    if (!state.NoteType) {
      errors["NoteType"] = "Note Type";
      setNotetypeBorderColor("#f31700");
    } else {
      setNotetypeBorderColor("");
    }

    if (showTypeOfFinNote && !state.TypeOfFinNote) {
      errors["TypeOfFinNote"] = "Type of Financial Note";
      setTypeOfFinBorderColor("#f31700");
    } else {
      setTypeOfFinBorderColor("");
    }

    if (showAmount && !amount) {
      errors["Amount"] = "Amount";
      setAmountBorderColor("#f31700");
    } else {
      setAmountBorderColor("");
    }

    // Bug fix - 297 - 27/03
    if (showAmount && Number(amount) <= 0) {
      errors["Amount"] = "Amount";
      setAmountBorderColor("#f31700");
    } else {
      setAmountBorderColor("");
    }

    if (!state.Search.trim()) {
      errors["Search"] = "Search Text";
      setSearchBorderColor("#f31700");
    } else {
      setSearchBorderColor("");
    }

    if (showPurpose && !purpose.trim()) {
      errors["Purpose"] = "Purpose";
      setPurposeBorderColor("#f31700");
    } else {
      setPurposeBorderColor("");
    }

    if (showPurpose1 && !purpose) {
      errors["Purpose1"] = "Purpose";
      setPurpose1BorderColor("#f31700");
    } else {
      setPurpose1BorderColor("");
    }

    if (showPurpose2 && !purpose) {
      errors["Purpose2"] = "Purpose";
      setPurpose2BorderColor("#f31700");
    } else {
      setPurpose2BorderColor("");
    }

    if (showPurpose1 && purpose === "Others" && !purposeOthers.trim()) {
      errors["Purpose2"] = "Others";
      setPurposeOthersBorderColor("#f31700");
    } else {
      setPurposeOthersBorderColor("");
    }

    if (noteapproverData.length === 0) {
      errors["Approvers"] = "Please select at least one Approver to submit request";
    }

    if (isApproverValid || isReviewerValid) {
      errors["Login user"] = "Login user cannot be part of Approvers/Reviewers.";
    }

    // === PDF and Word validation ===
    if (
      !wordandPdfWarring?.PDFInfo.fileName ||
      !wordandPdfWarring?.PDFInfo.isValid
    ) {
      errors["PDFInfo"] = "Please select Valid Pdf File";
    }

    if (
      noteapproverData.length > 0 &&
      isSecretaryExist &&
      (!wordandPdfWarring?.wordInfo.fileName ||
        !wordandPdfWarring?.wordInfo.isValid)
    ) {
      errors["wordDocInfo"] = "Please select Valid Word Doc File";
    }

    // === Supporting document validation ===
    if (Object.keys(supportDocWarning).length > 0 || totalFileSize > 26214400) {
      if (Object.keys(supportDocWarning).length > 0) {
        errors["Supporting Documents"] = "Please select valid files";
      }
      if (totalFileSize > 26214400) {
        errors["Supporting DocumentsMaxSize"] =
          "Cumulative size of all the supporting documents should not exceed 25 MB.";
      }
    }

    // === Final Approver designation validation ===
    // Final Approver designation validation (improved)
    // Validate that final approver designation matches NoteTo
    // if (noteapproverData && noteapproverData.length > 0 && state.NoteTo) {
    //   const finalApprover = noteapproverData[noteapproverData.length - 1]; // last approver
    //   const finalApproverDesignation = finalApprover.designation?.trim().toLowerCase();
    //   const noteToDesignation = state.NoteTo?.trim().toLowerCase();

    //   // Check if same designation already exists in the list
    //   const sameNoteToExists = noteapproverData.some(
    //     (approver) => approver.designation?.trim().toLowerCase() === noteToDesignation
    //   );

    //   // Show error if mismatch or not present
    //   if (!sameNoteToExists || finalApproverDesignation !== noteToDesignation) {
    //     errors["DesignationMismatch"] = "Please select designation same as Note To Field.";
    //   }
    // }


    // === Show validation errors ===
    if (Object.keys(errors).length > 0) {
      const errorMessagesArray = Object.values(errors);
      setErrorMessages(errorMessagesArray);
      setValidationErrors(true);
    }

    return errors;
  };


  //It helps to cancel the note 
  const handleAutoSave = async () => {
    let totalFileSize = 0;
    const inValidFileNames = Object.keys(supportDocWarning).filter(
      (fileName) => supportDocWarning[fileName]?.isValid === false
    );

    // Filter valid files
    const vaildMultiplefile = filesInfo.filter(obj =>
      !inValidFileNames.includes(obj.supportingDocumentFileName)
    );

    vaildMultiplefile.forEach(obj =>
      totalFileSize += obj.supportingDocumentPathLength
    );

    const fileupdated = filesInfo?.map(item => {
      const { fileObject, supportingDocumentPathLength: _, ...rest } = item;
      return rest;
    });

    try {
      const accessToken = await getAccessToken({ ...loginRequest, account }, instance);
      const updatedNoteReviewerData = notereviewerData
        .map((item, ind) => ({
          ...item,
          strApproverStatus: ind === 0 ? "Pending" : "Waiting",
        }))
        .map((item) =>
          Object.fromEntries(
            Object.entries(item).filter(([_, value]) => value !== null && value !== "")
          )
        );

      const updatedNoteApproverData = noteapproverData
        .map((item, ind) => ({
          ...item,
          strApproverStatus:
            notereviewerData.length === 0 && ind === 0 ? "Pending" : "Waiting",
        }))
        .map((item) =>
          Object.fromEntries(
            Object.entries(item).filter(([_, value]) => value !== null && value !== "")
          )
        );
      const base64PDFParams = wordandPdfWarring.PDFInfo.filePath;
      const base64WordParams = wordandPdfWarring.wordInfo.filePath || {};
      const cleanApproverData = (data) => data.map(({ isApproved, isRejected, workflowLogId, ...rest }) => rest);

      // Create the params object
      let params = {
        noteId: draftId || 0,
        noteFor: enumsObj.noteFor.find(x => x.dValue === "Regular").id,
        status: enumsObj.NoteStatus.find(x => x.dValue === "Draft").id,
        // departmentId: userDepartment ? departmentList.find(obj => obj.departmentName === userDepartment).departmentId : 0,
        departmentName: userDepartment,
        noteTo: state.NoteTo ? noteto.find(x => x.dValue === state.NoteTo).id : null,
        subject: state.Subject,
        designation: userDesignation,
        natureofNote: state.NatureOfNote ? natureofnote.find(x => x.dValue === state.NatureOfNote).id : 0,
        natureOfApprovalOrSanction: state.NatureOfApprSanc ? natureofapprsanc.find(x => x.dValue === state.NatureOfApprSanc).id : 0,
        noteType: state.NoteType ? notetype.find(x => x.dValue === state.NoteType).id : 0,
        financialType: state.TypeOfFinNote ? fintype.find(x => x.dValue === state.TypeOfFinNote).id : 0,
        amount: amount ? amount : 0,
        createdBy: accounts[0].username,
        modifiedBy: accounts[0].username,
        noteSupportingDocumentsDTO: totalFileSize > 26214400 ? [] : fileupdated,
        noteApproversDTO: cleanApproverData([
          ...updatedNoteReviewerData,
          ...updatedNoteApproverData
        ]),
        autosave: true,
        searchKeyword: state.Search,
        Purpose: (showPurpose1 && purpose === "Others") ? purposeOthers : purpose,
        NotePdfPathPart1: base64PDFParams.part1,
        NotePdfPathPart2: base64PDFParams.part2,
        NotePdfPathPart3: base64PDFParams.part3,
        NotePdfPathPart4: base64PDFParams.part4,
        NotePdfPathPart5: base64PDFParams.part5,
        NotePdfPathPart6: base64PDFParams.part6,
        NotePdfPathPart7: base64PDFParams.part7,
        NotePdfPathPart8: base64PDFParams.part8,
        NotePdfPathPart9: base64PDFParams.part9,
        NotePdfPathPart10: base64PDFParams.part10,
        NoteWordPathPart1: base64WordParams.part1,
        NoteWordPathPart2: base64WordParams.part2,
        NoteWordPathPart3: base64WordParams.part3,
        NoteWordPathPart4: base64WordParams.part4,
        NoteWordPathPart5: base64WordParams.part5,
        NoteWordPathPart6: base64WordParams.part6,
        NoteWordPathPart7: base64WordParams.part7,
        NoteWordPathPart8: base64WordParams.part8,
        NoteWordPathPart9: base64WordParams.part9,
        NoteWordPathPart10: base64WordParams.part10,
        NotePdfFileName: (wordandPdfWarring?.PDFInfo.fileName && wordandPdfWarring?.PDFInfo.isValid === false) ? null : wordandPdfWarring.PDFInfo.fileName,
        NoteWordFileName: (isSecretaryExist && (wordandPdfWarring?.wordInfo.fileName && wordandPdfWarring?.wordInfo.isValid === false)) ? null : wordandPdfWarring.wordInfo.fileName || null,
        isConfidential: isConfidential ? true : false,
      };

      if (!params.financialType) {
        params.financialType = 0;
      }
      if (!params.status) {
        params.status = 0;
      }
      if (!params.natureofNote) {
        params.natureofNote = 0;
      }
      if (!params.natureOfApprovalOrSanction) {
        params.natureOfApprovalOrSanction = 0;
      }
      if (!params.noteType) {
        params.noteType = 0;
      }

      if (!params.amount) {
        params.amount = 0;
      }

      // Remove other fields that are null or empty strings
      params = Object.fromEntries(
        Object.entries(params).filter(
          ([key, value]) =>
            value !== null && value !== "" && (value !== 0 || ["financialType", "status", "natureofNote", "noteType", "natureOfApprovalOrSanction", "amount"].includes(key))
        )
      );

      await fetch(
        `${API_BASE_URL}${draftId === 0 ? API_ENDPOINTS.eNote_AddNote : API_ENDPOINTS.eNote_EditNote}`,
        {
          method: "POST",
          headers: { ...API_COMMON_HEADERS, Authorization: `Bearer ${accessToken}` },
          body: JSON.stringify(params),
        }
      ).then(async (data) => {
        const res = await data.json();
        if (res.statusMessage === "Success") {
          setDraftId(res.noteId);
        } else {
          console.log("Something went wrong, please try again.");
        }
        setIsLoading(false);
      }).catch((err) => {
        console.log(err, "error");
      });
    } catch (error) {
      console.log("Error", error);
    }
  };

  const onCancelNote = async () => {
    setIsLoading(true);
    setVisibleCancelCfrm(false);
    let params = {
      noteId: noteId,
      createdBy: accounts[0].username,
    };
    const accessToken = await getAccessToken({ ...loginRequest, account }, instance);
    await fetch(`${API_BASE_URL}${API_ENDPOINTS.eNote_CancelNote}`, {
      method: "POST",
      headers: { ...API_COMMON_HEADERS, Authorization: `Bearer ${accessToken}` },
      body: JSON.stringify(params),
    })
      .then(async (data) => {
        const res = await data.json();
        if (res.statusMessage === "Success") {
          setSuccessMessage("The request for cancellation has been successfull.");
          setVisibleSave(!visiblesave);
        } else {
          setShowNotification(true);
          setNotificationMsg("Something went wrong please try again.");
        }
        setIsLoading(false);
      })
      .catch((err) => {
        setShowNotification(true);
        setNotificationMsg("Something went wrong please try again.");
        console.log(err);
        setIsLoading(false);
      });
  };

  /* it helps tp save or update draft data on Save as draft*/
  const handleSave = async () => {
    const errors = {};
    setIsLoading(true);
    setisSaveAndSubmitActionCompleted(true);
    let totalFileSize = 0;
    let isFilesValid = true;

    filesInfo.map((x) => {
      isFilesValid = !(x.supportingDocumentFileName in supportDocWarning);
    });

    const inValidFileNames = Object.keys(supportDocWarning).filter(
      (fileName) => supportDocWarning[fileName]?.isValid === false
    );

    if (filesInfo.length > 0) {
      const vaildMultiplefile = filesInfo.filter(obj => {
        if (!inValidFileNames.includes(obj.supportingDocumentFileName)) {
          return obj;
        }
      });
      vaildMultiplefile.map(obj =>
        totalFileSize = totalFileSize + obj.supportingDocumentPathLength
      );
    }

    let fileupdated = filesInfo?.map(item => {
      const { supportingDocumentPathLength: _, ...rest } = item;
      return rest;
    });

    // 🔹 File validation
    if (isSecretaryExist && (wordandPdfWarring?.wordInfo.fileName && wordandPdfWarring?.wordInfo.isValid === false)) {
      errors["wordDocInfo"] = "Please select vaild file";
    }
    if (wordandPdfWarring?.PDFInfo.fileName && wordandPdfWarring?.PDFInfo.isValid === false) {
      errors["PDFInfo"] = "Please select vaild file";
    }
    if ((filesInfo.length > 0 && !isFilesValid) || totalFileSize > 26214400) {
      if (filesInfo.length > 0 && !isFilesValid) {
        errors["Supporting Documents"] = "Please select vaild files";
      }
      if (totalFileSize > 26214400) {
        errors["Supporting DocumentsMaxSize"] = "Cumulative size of all the supporting documents should not exceed 25 MB.";
      }
    }

    // 🔹 NEW VALIDATION: Final approver's designation must match NoteTo
    // if (noteapproverData && noteapproverData.length > 0) {
    //   const finalApprover = noteapproverData[noteapproverData.length - 1];
    //   const noteToDesignation = state.NoteTo; // Assuming NoteTo stores the designation text

    //   if (finalApprover.designation !== noteToDesignation) {
    //     errors["DesignationMismatch"] = "Please select designation same as Note To Field.";
    //   }
    // }

    // ✅ Proceed if no validation errors
    if (Object.keys(errors).length === 0) {
      try {
        const updatedNoteReviewerData = notereviewerData
          .map((item, ind) => ({
            ...item,
            strApproverStatus: ind === 0 ? "Pending" : "Waiting",
          }))
          .map((item) =>
            Object.fromEntries(
              Object.entries(item).filter(([_, value]) => value !== null && value !== "")
            )
          );

        const updatedNoteApproverData = noteapproverData
          .map((item, ind) => ({
            ...item,
            strApproverStatus:
              notereviewerData.length === 0 && ind === 0 ? "Pending" : "Waiting",
          }))
          .map((item) =>
            Object.fromEntries(
              Object.entries(item).filter(([_, value]) => value !== null && value !== "")
            )
          );

        const base64PDFParams = wordandPdfWarring.PDFInfo.filePath;
        const base64WordParams = wordandPdfWarring.wordInfo.filePath || {};
        const cleanApproverData = (data) =>
          data.map(({ isApproved, isRejected, workflowLogId, ...rest }) => rest);

        let params = {
          noteId: isNewForm ? draftId : noteId,
          noteFor: enumsObj.noteFor.find(x => x.dValue === "Regular").id,
          status: enumsObj.NoteStatus.find(x => x.dValue === "Draft").id,
          departmentName: userDepartment,
          noteTo: state.NoteTo ? noteto.find(x => (x.dValue === state.NoteTo)).id : null,
          subject: state.Subject,
          designation: userDesignation,
          natureofNote: state.NatureOfNote ? natureofnote.find(x => (x.dValue === state.NatureOfNote)).id : 0,
          natureOfApprovalOrSanction: state.NatureOfApprSanc ? natureofapprsanc.find(x => (x.dValue === state.NatureOfApprSanc)).id : 0,
          noteType: state.NoteType ? notetype.find(x => (x.dValue === state.NoteType)).id : 0,
          financialType: state.TypeOfFinNote ? fintype.find(x => (x.dValue === state.TypeOfFinNote)).id : 0,
          amount: amount ? amount : 0,
          createdBy: accounts[0].username,
          modifiedBy: accounts[0].username,
          noteSupportingDocumentsDTO: fileupdated,
          noteApproversDTO: cleanApproverData([
            ...updatedNoteReviewerData,
            ...updatedNoteApproverData
          ]),
          autosave: false,
          searchKeyword: state.Search,
          Purpose: (showPurpose1 && purpose === "Others") ? purposeOthers : purpose,
          NotePdfPath: isNewForm ? null : (wordandPdfWarring.PDFInfo.filePath !== notepath ? null : notepath),
          NotePdfPathPart1: base64PDFParams.part1,
          NotePdfPathPart2: base64PDFParams.part2,
          NotePdfPathPart3: base64PDFParams.part3,
          NotePdfPathPart4: base64PDFParams.part4,
          NotePdfPathPart5: base64PDFParams.part5,
          NotePdfPathPart6: base64PDFParams.part6,
          NotePdfPathPart7: base64PDFParams.part7,
          NotePdfPathPart8: base64PDFParams.part8,
          NotePdfPathPart9: base64PDFParams.part9,
          NotePdfPathPart10: base64PDFParams.part10,
          NotePdfFileName: wordandPdfWarring.PDFInfo.fileName,
          NoteWordPathPart1: base64WordParams.part1,
          NoteWordPathPart2: base64WordParams.part2,
          NoteWordPathPart3: base64WordParams.part3,
          NoteWordPathPart4: base64WordParams.part4,
          NoteWordPathPart5: base64WordParams.part5,
          NoteWordPathPart6: base64WordParams.part6,
          NoteWordPathPart7: base64WordParams.part7,
          NoteWordPathPart8: base64WordParams.part8,
          NoteWordPathPart9: base64WordParams.part9,
          NoteWordPathPart10: base64WordParams.part10,
          NoteWordPath: isNewForm ? null : (wordandPdfWarring.wordInfo.filePath !== wordPath ? null : wordPath),
          NoteWordFileName: wordandPdfWarring.wordInfo.fileName,
          isConfidential: isConfidential ? true : false,
        };

        // Default fallbacks for missing enum IDs
        if (!params.financialType) params.financialType = 0;
        if (!params.status) params.status = 0;
        if (!params.natureofNote) params.natureofNote = 0;
        if (!params.natureOfApprovalOrSanction) params.natureOfApprovalOrSanction = 0;
        if (!params.noteType) params.noteType = 0;
        if (!params.amount) params.amount = 0;

        // Remove null/empty fields except key numeric fields
        params = Object.fromEntries(
          Object.entries(params).filter(
            ([key, value]) =>
              value !== null && value !== "" &&
              (value !== 0 || ["financialType", "status", "natureofNote", "noteType", "natureOfApprovalOrSanction", "amount"].includes(key))
          )
        );

        const accessToken = await getAccessToken({ ...loginRequest, account }, instance);
        await fetch(
          `${API_BASE_URL}${draftId === 0 ? API_ENDPOINTS.eNote_AddNote : API_ENDPOINTS.eNote_EditNote}`,
          {
            method: "POST",
            headers: { ...API_COMMON_HEADERS, Authorization: `Bearer ${accessToken}` },
            body: JSON.stringify(params),
          }
        )
          .then(response => response.json())
          .then(data => {
            if (data.statusMessage === "Success") {
              setSuccessMessage("The request for eNote has been drafted successfully.");
              setVisibleSave(true);
            } else {
              setShowNotification(true);
              setNotificationMsg(`eNote form submission failed.\n${data.statusMessage}`);
            }
          })
          .catch(err => {
            console.error("Error submitting form:", err);
            setShowNotification(true);
            setNotificationMsg(err);
          });
      } catch (error) {
        console.error("Error submitting form:", error);
        setShowNotification(true);
        setNotificationMsg(error);
      }
    } else {
      // 🔹 Show specific error message
      setShowNotification(true);
      const firstError = Object.values(errors)[0];
      setNotificationMsg(firstError || "Invalid files attached. Kindly remove the invalid files.");
    }

    setIsLoading(false);
  };


  //Submit handler helps to vlidate mandatory fields and prompt confirm dailog
  const handleSubmit = async () => {
    setIsLoading(true);
    setVisible(false);
    setisSaveAndSubmitActionCompleted(true);

    let fileupdated = filesInfo?.map(item => {
      const { supportingDocumentPathLength: _, ...rest } = item; // Destructure 'supportingDocumentPathLength' and collect the rest
      // Filter out null or empty string values from the rest of the object
      return Object.fromEntries(
        Object.entries(rest).filter(([_, value]) => value !== null && value !== "")
      );
    });
    try {
      const updatedNoteReviewerData = notereviewerData
        .map((item, ind) => ({
          ...item,
          strApproverStatus: ind === 0 ? "Pending" : "Waiting",
        }))
        .map((item) =>
          Object.fromEntries(
            Object.entries(item).filter(([_, value]) => value !== null && value !== "")
          )
        );

      const updatedNoteApproverData = noteapproverData
        .map((item, ind) => ({
          ...item,
          strApproverStatus:
            notereviewerData.length === 0 && ind === 0 ? "Pending" : "Waiting",
        }))
        .map((item) =>
          Object.fromEntries(
            Object.entries(item).filter(([_, value]) => value !== null && value !== "")
          )
        );
      // Added for base 64 split params ---> Kavya(18-07)
      const base64PDFParams = wordandPdfWarring.PDFInfo.filePath || {};
      const base64WordParams = wordandPdfWarring.wordInfo.filePath || {};
      const cleanApproverData = (data) => data.map(({ isApproved, isRejected, workflowLogId, ...rest }) => rest);

      let params = {
        noteId: isNewForm ? draftId : noteId,
        noteFor: enumsObj.noteFor.find(x => x.dValue === "Regular").id,
        status: enumsObj.NoteStatus.find(x => x.dValue === "Submitted").id,
        departmentName: userDepartment,
        noteTo: state.NoteTo ? noteto.find(x => (x.dValue === state.NoteTo)).id : 0,
        subject: state.Subject,
        designation: userDesignation,
        natureofNote: state.NatureOfNote ? natureofnote.find(x => (x.dValue === state.NatureOfNote)).id : 0,
        natureOfApprovalOrSanction: state.NatureOfApprSanc ? natureofapprsanc.find(x => (x.dValue === state.NatureOfApprSanc)).id : 0,
        noteType: state.NoteType ? notetype.find(x => (x.dValue === state.NoteType)).id : 0,
        financialType: state.TypeOfFinNote ? fintype.find(x => (x.dValue === state.TypeOfFinNote)).id : 0,
        amount: amount ? amount : 0,
        createdBy: accounts[0].username,
        modifiedBy: accounts[0].username,
        noteSupportingDocumentsDTO: fileupdated,// filesInfo,
        noteApproversDTO: cleanApproverData([
          ...updatedNoteReviewerData,
          ...updatedNoteApproverData
        ]),
        autosave: false,
        searchKeyword: state.Search,
        Purpose: (showPurpose1 && purpose === "Others") ? purposeOthers : purpose,
        // NotePdfPath: wordandPdfWarring.PDFInfo.filePath,
        NotePdfPath: isNewForm ? null : (wordandPdfWarring.PDFInfo.filePath !== notepath ? null : notepath),
        // Added for base 64 split params ---> Kavya(18-07)
        NotePdfPathPart1: base64PDFParams.part1,
        NotePdfPathPart2: base64PDFParams.part2,
        NotePdfPathPart3: base64PDFParams.part3,
        NotePdfPathPart4: base64PDFParams.part4,
        NotePdfPathPart5: base64PDFParams.part5,
        NotePdfPathPart6: base64PDFParams.part6,
        NotePdfPathPart7: base64PDFParams.part7,
        NotePdfPathPart8: base64PDFParams.part8,
        NotePdfPathPart9: base64PDFParams.part9,
        NotePdfPathPart10: base64PDFParams.part10,
        NotePdfFileName: wordandPdfWarring.PDFInfo.fileName,
        NoteWordPathPart1: base64WordParams.part1,
        NoteWordPathPart2: base64WordParams.part2,
        NoteWordPathPart3: base64WordParams.part3,
        NoteWordPathPart4: base64WordParams.part4,
        NoteWordPathPart5: base64WordParams.part5,
        NoteWordPathPart6: base64WordParams.part6,
        NoteWordPathPart7: base64WordParams.part7,
        NoteWordPathPart8: base64WordParams.part8,
        NoteWordPathPart9: base64WordParams.part9,
        NoteWordPathPart10: base64WordParams.part10,
        NoteWordPath: isNewForm ? null : (wordandPdfWarring.wordInfo.filePath !== wordPath ? null : wordPath) || null,
        NoteWordFileName: wordandPdfWarring.wordInfo.fileName || null,
        isConfidential: isConfidential ? true : false,
      };

      if (!params.financialType) {
        params.financialType = 0;
      }
      if (!params.status) {
        params.status = 0;
      }
      if (!params.natureofNote) {
        params.natureofNote = 0;
      }
      if (!params.natureOfApprovalOrSanction) {
        params.natureOfApprovalOrSanction = 0;
      }
      if (!params.noteType) {
        params.noteType = 0;
      }

      if (!params.amount) {
        params.amount = 0;
      }

      // Remove other fields that are null or empty strings
      params = Object.fromEntries(
        Object.entries(params).filter(
          ([key, value]) =>
            value !== null && value !== "" && (value !== 0 || ["financialType", "status", "natureofNote", "noteType", "natureOfApprovalOrSanction", "amount"].includes(key))
        )
      );
      const accessToken = await getAccessToken({ ...loginRequest, account }, instance);
      await fetch(
        `${API_BASE_URL}${(isNewForm && draftId === 0) ? API_ENDPOINTS.eNote_AddNote : API_ENDPOINTS.eNote_EditNote}`,
        {
          method: "POST",
          headers: { ...API_COMMON_HEADERS, Authorization: `Bearer ${accessToken}` },
          body: JSON.stringify(params),
        }
      ).then(response => {
        return response.json();
      }).then(data => {
        if (data.statusMessage === "Success") {
          setAlertVisible(true);
        } else {
          setShowNotification(true);
          setNotificationMsg(`eNote form submission failed. 
            ${data.statusMessage}`);
        }
      }).catch(err => {
        console.error(err, "Error submitting form");
        setShowNotification(true);
        setNotificationMsg(err);
      })
    } catch (error) {
      console.error("Error submitting form:", error);
      setShowNotification(true);
      setNotificationMsg(error);
    }
    setIsLoading(false);
  };

  const handleOpenDialog = async () => {

    let isApproverValid = noteapproverData.find((x) => x.approverEmail === accounts[0].username) !== undefined;
    let isReviewerValid = notereviewerData.find((x) => x.approverEmail === accounts[0].username) !== undefined;

    if (!userDepartment) {
      setShowNotification(true);
      setNotificationMsg("User department cannot be blank. Please contact system administrator.");
      return;
    }
    const errors = validateForm(isApproverValid, isReviewerValid);
    if (Object.keys(errors).length === 0) {
      setVisible(!visible);
    }
  };

  // handle close  dialog
  const handleCloseDialog = async () => {
    setVisible(!visible);
  };

  // handle save close 
  const handleSaveClose = async () => {
    setVisibleSave(false);
  };

  // handle close validation dialog
  const handleClosevalidationDialog = async () => {
    setValidationErrors(false);
  };

  //It helps to validate and convert to Base64 to attach documents - Word Info
  const convertBase64Word = () => {
    var selectedFile = document.getElementById("WordDocfile").files;
    let cstWarningMsg = "";
    let isValidFile = true;
    if (selectedFile.length > 0) {
      var fileExtension = selectedFile[0].name.split(".");
      const fileName = selectedFile[0].name

      if (
        !(
          fileName.toLowerCase().endsWith("docx") ||
          fileName.toLowerCase().endsWith("doc" || "DOC")
        )
      ) {
        cstWarningMsg = "File type not allowed";
        isValidFile = false;
      }

      if (selectedFile[0].size > 10485760) {
        cstWarningMsg = "File size should not exceed more then 10 MB";
        isValidFile = false;
      }
      if (checkingSpl(selectedFile[0].name)) {
        isValidFile = false;
        cstWarningMsg = "File name sholud not contain special characters";
      }

      if (isValidFile) {
        var fileToLoad = selectedFile[0];
        // FileReader function for read the file.
        var fileReader = new FileReader();
        // Onload of file read the file content
        fileReader.onload = function (fileLoadedEvent) {
          let base64 = fileLoadedEvent.target.result;
          const RemovedHeaderFormBase64 = base64.split(",")[1];

          const partLength = Math.ceil(RemovedHeaderFormBase64.length / 10);
          const parts = [];

          for (let i = 0; i < 10; i++) {
            parts.push(RemovedHeaderFormBase64.slice(i * partLength, (i + 1) * partLength));
          }
          SetWordPDFInfowarring({
            ...wordandPdfWarring,
            wordInfo: {
              fileExtension: `.${fileExtension[fileExtension.length - 1]}`,
              fileName: selectedFile[0].name,
              filePath: {
                part1: parts[0],
                part2: parts[1],
                part3: parts[2],
                part4: parts[3],
                part5: parts[4],
                part6: parts[5],
                part7: parts[6],
                part8: parts[7],
                part9: parts[8],
                part10: parts[9]
              },
              isValid: isValidFile,
              warningMsg: cstWarningMsg
            },
          });
        };
        // Convert data to base64
        fileReader.readAsDataURL(fileToLoad);
      } else {
        SetWordPDFInfowarring({
          ...wordandPdfWarring,
          wordInfo: {
            fileExtension: `.${fileExtension[fileExtension.length - 1]}`,
            fileName: selectedFile[0].name,
            filePath: null,
            isValid: isValidFile,
            warningMsg: cstWarningMsg
          },
        });
      }
    }
  };

  // Added for base 64 split params ---> Kavya(18-07)
  const convertPDFToBase64 = () => {
    var selectedFile = document.getElementById("PDFDocfile").files;
    let cstWarningMsg = "";
    let isValidFile = true;

    if (selectedFile.length > 0) {
      var fileExtension = selectedFile[0].name.split(".");
      const fileName = selectedFile[0].name;

      if (!(fileName.toLowerCase().endsWith("pdf"))) {
        cstWarningMsg = "File type not allowed";
        isValidFile = false;
      }

      if (selectedFile[0].size > 26214400) {
        cstWarningMsg = "File size should not exceed more than 25 MB";
        isValidFile = false;
      }

      if (checkingSpl(selectedFile[0].name)) {
        isValidFile = false;
        cstWarningMsg = "File name should not contain special characters";
      }

      if (isValidFile) {
        var fileToLoad = selectedFile[0];
        var fileReader = new FileReader();

        fileReader.onload = function (fileLoadedEvent) {
          let base64 = fileLoadedEvent.target.result;
          const RemovedHeaderFormBase64 = base64.split(",")[1];

          const partLength = Math.ceil(RemovedHeaderFormBase64.length / 10);
          const parts = [];

          for (let i = 0; i < 10; i++) {
            parts.push(RemovedHeaderFormBase64.slice(i * partLength, (i + 1) * partLength));
          }

          SetWordPDFInfowarring({
            ...wordandPdfWarring,
            PDFInfo: {
              fileExtension: `.${fileExtension[fileExtension.length - 1]}`,
              fileName: selectedFile[0].name,
              filePath: {
                part1: parts[0],
                part2: parts[1],
                part3: parts[2],
                part4: parts[3],
                part5: parts[4],
                part6: parts[5],
                part7: parts[6],
                part8: parts[7],
                part9: parts[8],
                part10: parts[9]
              },
              isValid: isValidFile,
              warningMsg: cstWarningMsg
            }
          });
        };

        fileReader.readAsDataURL(fileToLoad);
      } else {
        SetWordPDFInfowarring({
          ...wordandPdfWarring,
          PDFInfo: {
            fileExtension: `.${fileExtension[fileExtension.length - 1]}`,
            fileName: selectedFile[0].name,
            filePath: {},
            isValid: isValidFile,
            warningMsg: cstWarningMsg
          }
        });
      }
    }
  };

  //Helps to validate multiple docs - Support documents
  const checkingSpl = (test) => {
    var specialCharPattern = /[!@#$%^&*(){}\[\];:,<>\?\/\\]/;

    // Test the input string against the pattern
    return specialCharPattern.test(test);
  };

  //It helps to validate and attach documents - support documents

  // base64 code commented
  const multipleDocUpload = () => {
    let isValidFile = true;
    let waringmsg = "";
    let reamingFileSize = 26214400;
    const inValidFileNames = Object.keys(supportDocWarning).filter(
      (fileName) => supportDocWarning[fileName]?.isValid === false
    );

    if (filesInfo.length > 0) {
      const vaildMultiplefile = filesInfo.filter(obj => {
        if (!inValidFileNames.includes(obj.supportingDocumentFileName)) {
          return obj;
        }
      });
      vaildMultiplefile.forEach(obj =>
        reamingFileSize -= obj.supportingDocumentPathLength
      );
    }

    const arrayExtension = [
      ".pdf", ".doc", ".docx", ".xlsx",
      ".PDF", ".DOC", ".DOCX", ".XLSX"
    ];
    const validname = /[!@#$%^&*(){}\[\];:,<>\?\/\\]/;
    const selectedFile = document.getElementById("multiDoc").files;
    const TempfileInfo = filesInfo;

    const promises = [];
    const fileCount = selectedFile.length;

    for (let i = 0; i < fileCount; i++) {
      const fileToLoad = selectedFile[i];
      const fileExtession = fileToLoad.name.split(".");

      if (!filesInfo?.some(obj => obj.supportingDocumentFileName === fileToLoad.name)) {
        if (!arrayExtension.includes(`.${fileExtession[fileExtession.length - 1]}`)) {
          isValidFile = false;
          waringmsg = "File type is not allowed";
          SetSupportDocWarning({
            ...supportDocWarning,
            [fileToLoad.name]: {
              filename: fileToLoad.name,
              isValid: true,
              warningMsg: waringmsg,
              fileExtnsion: fileExtession[fileExtession.length - 1],
            },
          });
        } else if (validname.test(fileToLoad.name)) {
          // isValidFile = false;
          waringmsg = "File name should not contain special characters";
          SetSupportDocWarning({
            ...supportDocWarning,
            [fileToLoad.name]: {
              filename: fileToLoad.name,
              isValid: true,
              warningMsg: waringmsg,
              fileExtnsion: fileExtession[fileExtession.length - 1],
            },
          });
        } else if (fileToLoad.size > reamingFileSize) {
          // isValidFile = false;
          waringmsg = "Cumulative size of all the supporting documents should not be exceeded 25 MB.";
          setSupportingError(waringmsg);
        }
      }

      const fileReader = new FileReader();

      const promise = new Promise((resolve) => {
        fileReader.onload = function (fileLoadedEvent) {
          // Added for base 64 split ---> Kavya(18-07)
          const base64 = fileLoadedEvent.target.result.split(",")[1];
          const partLength = Math.ceil(base64.length / 10);
          const parts = [];
          for (let j = 0; j < 10; j++) {
            parts.push(base64.slice(j * partLength, (j + 1) * partLength));
          }

          resolve({
            name: fileToLoad.name,
            parts: parts,
            size: fileToLoad.size
          });
        };
      });

      fileReader.readAsDataURL(fileToLoad);
      promises.push(promise);
    }

    Promise.all(promises)
      .then((fileDataArray) => {
        const updatedTempFileInfo = fileDataArray.reduce((acc, fileData) => {
          const ObjExist = acc.map((obj) => obj.supportingDocumentFileName);
          if (!ObjExist.includes(fileData.name)) {
            acc.push({
              noteSupportingDocumentId: 0,
              noteId: 0,
              supportingDocumentFileName: fileData.name,
              createdDate: new Date(),
              createdBy: accounts[0].username,
              modifiedDate: new Date(),
              modifiedBy: accounts[0].username,
              // Added for base 64 split params parts ---> Kavya(18-07)
              supportingDocumentPathPart1: fileData.parts[0],
              supportingDocumentPathPart2: fileData.parts[1],
              supportingDocumentPathPart3: fileData.parts[2],
              supportingDocumentPathPart4: fileData.parts[3],
              supportingDocumentPathPart5: fileData.parts[4],
              supportingDocumentPathPart6: fileData.parts[5],
              supportingDocumentPathPart7: fileData.parts[6],
              supportingDocumentPathPart8: fileData.parts[7],
              supportingDocumentPathPart9: fileData.parts[8],
              supportingDocumentPathPart10: fileData.parts[9],
              supportingDocumentPathLength: fileData.size
            });
          }
          return acc;
        }, [...TempfileInfo]);

        SetFileinfo(updatedTempFileInfo);
      })
      .catch((error) => {
        console.error("Error reading files:", error);
      });
  };

  // handle Reviewer chaange
  const handleComboChangeReviewer = (event) => {
    setSelectedReviewer(event.value);
  };

  // handle  Approver change
  const handleComboChangeApprover = (event) => {
    setComboValueApprover(event.value);
  };

  //Helps to add selected reviewer to the reviewers table
  const handleAddRow = async () => {
    if (!selectedReveiwer) {
      setShowNotification(true);
      setNotificationMsg("Please select reviewer then click on Add.");
      return;
    }

    // Added restricted designation for MD & CEO and Executive Director for adding reviewer- 08/11
    const restrictedDesignations = ["MD & CEO", "ED"];
    if (restrictedDesignations.includes(selectedReveiwer.jobTitle)) {
      setShowNotification(true);
      setNotificationMsg("MD / ED can only be selected as Approver!");
      return;
    }

    let isApproverValid =
      noteapproverData.find((x) => x.approverEmail === selectedReveiwer.userPrincipalName) ===
      undefined;
    let isReviewerValid =
      notereviewerData.find((x) => x.approverEmail === selectedReveiwer.userPrincipalName) ===
      undefined;

    if (isApproverValid && isReviewerValid && selectedReveiwer.userPrincipalName !== accounts[0].username && selectedReveiwer.userPrincipalName !== state.createdBy) {
      const newRow = {
        approverType: 1,
        approverEmail: selectedReveiwer.userPrincipalName,
        approverOrder: notereviewerData.length + 1,
        approverStatus: 1,
        // Bug fix - 294 - 27/03
        srNo: selectedReveiwer.srNo,
        designation: selectedReveiwer.jobTitle,
        approverEmailName: selectedReveiwer.displayName,
        createdDate: new Date(),
        createdBy: accounts[0].username,
        modifiedDate: new Date(),
        modifiedBy: accounts[0].username,
      };
      setNoteReviewerData((prevData) => [...prevData, newRow]);

    } else {
      setShowNotification(true);
      setNotificationMsg(
        "The selected reviewer cannont be same as existing Reviewers/Requester/CurrentActioner."
      );
    }
    setSelectedReviewer(null);
    setOrgEmployees([]);
  };

  const handleOpenReveiwer = () => {
    if (!selectedReveiwer) {
      setOrgEmployees([]);
    } else {
      const filteredUsers = allOrgUsers.filter(user =>
        user.displayName.toLowerCase().includes(selectedReveiwer.displayName.toLowerCase())
      );
      setOrgEmployees(filteredUsers);
    }
  }

  const handleOpenApprover = () => {
    if (!combovalueApprover) {
      setOrgEmployees([]);
    } else {
      const filteredUsers = allOrgUsers.filter(user =>
        user.displayName.toLowerCase().includes(combovalueApprover.displayName.toLowerCase())
      );
      setOrgEmployees(filteredUsers);
    }
  }

  //Helps to add selected approver to the approvers table
  const handleAddRowApprover = async () => {
    if (!combovalueApprover) {
      setShowNotification(true);
      setNotificationMsg("Please select approver then click on Add.");
      return;
    }

    let isApproverValid =
      noteapproverData.find((x) => x.approverEmail === combovalueApprover.userPrincipalName) ===
      undefined;
    let isReviewerValid =
      notereviewerData.find((x) => x.approverEmail === combovalueApprover.userPrincipalName) ===
      undefined;

    if (isApproverValid && isReviewerValid && combovalueApprover.userPrincipalName !== accounts[0].username && combovalueApprover.userPrincipalName !== state.createdBy) {

      const newRow = {
        approverType: 2,
        approverEmail: combovalueApprover.userPrincipalName,
        approverOrder: noteapproverData.length + 1,
        approverStatus: 1,
        srNo: combovalueApprover.srNo,
        designation: combovalueApprover.jobTitle,
        approverEmailName: combovalueApprover.displayName,
        createdDate: new Date(),
        createdBy: accounts[0].username,
        modifiedDate: new Date(),
        modifiedBy: accounts[0].username
      };
      setNoteApproverData((prevData) => [...prevData, newRow]);

      let upApprObj = [...noteapproverData, newRow];
      if (upApprObj?.length > 0) {
        const secretaryValChk = await findSecrateries(upApprObj, userDepartment);
        if (!secretaryValChk) {
          SetWordPDFInfowarring({
            ...wordandPdfWarring,
            wordInfo: {
              fileExtension: "",
              fileName: "",
              filePath: "",
              isValid: false,
            },
          });
        }
      } else {
        SetWordPDFInfowarring({
          ...wordandPdfWarring,
          wordInfo: {
            fileExtension: "",
            fileName: "",
            filePath: "",
            isValid: false,
          },
        });
      }
    } else {
      setShowNotification(true);
      setNotificationMsg(
        "The selected approver cannont be same as existing Reviewers/Requester/CurrentActioner."
      );
    }
    setComboValueApprover(null);
    setOrgEmployees([]);
  };

  //It helps to remove Word/PDF attachments
  const onRemoveAttachmentWarning = (key) => {
    if (key === "PDFInfo") {
      SetWordPDFInfowarring({
        ...wordandPdfWarring,
        PDFInfo: {
          fileExtension: "",
          fileName: "",
          warningMsg: "",
          filePath: "",
          isValid: false,
        },
      });
    }
    if (key === "wordInfo") {
      SetWordPDFInfowarring({
        ...wordandPdfWarring,
        [key]: { fileExtension: "", fileName: "", isValid: false },
      });
    }
    setFileKey(Date.now());
  };

  //Helps to attache icon based on file extension
  const getFileIcon = (filename) => {
    const fileSplit = filename.split(".");
    const extnsion = `.${fileSplit[fileSplit.length - 1]}`;
    const fileExtension = extnsion.toLowerCase();
    let fileType = fileIcon;
    if (fileExtension === ".pdf") {
      fileType = filePdfIcon;
    }
    if (fileExtension === ".doc" || fileExtension === ".docx") {
      fileType = fileWordIcon;
    }
    if (fileExtension === ".png" || fileExtension === ".jpg") {
      fileType = fileImageIcon;
    }
    if (fileExtension === ".txt") {
      fileType = fileTxtIcon;
    }
    if (fileExtension === ".xlsx") {
      fileType = fileDataIcon;
    }
    return fileType;
  };

  // remove multiple attchemnts - support documents
  const onRemoveMultiAttachment = (id) => {
    const filename = filesInfo.find((obj, index) => index === id).supportingDocumentFileName;
    delete supportDocWarning[filename];
    SetSupportDocWarning(supportDocWarning);
    let totalFileSize = 0;
    const inValidFileNames = Object.keys(supportDocWarning).filter(
      (fileName) => supportDocWarning[fileName]?.isValid === false
    );
    if (filesInfo.length > 0) {
      const vaildMultiplefile = filesInfo.filter(obj => {
        if (!inValidFileNames.includes(obj.supportingDocumentFileName)) {
          return obj;
        }
      });
      vaildMultiplefile.map((obj, index) => {
        if (index !== id) {
          totalFileSize = totalFileSize + obj.supportingDocumentPathLength
        }
      });

      if (totalFileSize <= 26214400) {
        setSupportingError("");
      }
    }
    SetFileinfo(filesInfo.filter((obj, ind) => ind !== id));
    setFileKey(Date.now()); // Added to reset the file input
  };

  //Helps to filter the people picker dropdown vaules
  const pplFilterMultiColumn = async (event) => {

    if (event.filter.value.length >= 4) {
      setOrgEmployees([]);
      const accessToken = await getAccessToken({
        ...loginRequest,
        account
      }, instance);

      await fetch(
        `${API_BASE_URL}${API_ENDPOINTS.Search_UserDetails(
          event.filter.value
        )}`,
        {
          method: "GET",
          headers: {
            ...API_COMMON_HEADERS,
            Authorization: `Bearer ${accessToken}`,
          },
        }
      )
        .then((response) => {
          return response.json();
        })
        .then((data) => {
          const orgUsers = data.map(x => {
            return { ...x, department: x.department === null ? "NA" : x.department, displayName: x.displayName === null ? "NA" : x.displayName, jobTitle: x.jobTitle === null ? "NA" : x.jobTitle, userPrincipalName: x.userPrincipalName, srNo: x.srNo === null ? "NA" : x.srNo }
          });
          setOrgEmployees(orgUsers);
          setAllOrgUsers(orgUsers);
        })
        .catch((err) => {
          setOrgEmployees([]);
          console.log(err);
        });
    } else if (event.filter.value.length < 4) {
      setOrgEmployees([]);
    }
    else {
      setOrgEmployees([]);
    }
  };

  //HTML - render
  return (
    <div>
      <Navbar header="IB Smart Office - eNote" />
      <Sidebar />
      {/* Added top navigation for menu --> 23-09 */}
      <MenuNavContainer isMenuPage />
      {isVieweEligible === "false" ? <> <Unauthorized />  <div className="pgFooterContainer">
        <Footer />
      </div> </> : <>
        <div className="FormMaincontainer">
          <div className="container largeScreen">
            <div className="HeaderButtonsContainer row">
              <div className="cstformHdrLeftContainer mobileTitleNewForm">
                {isNewForm ? "" : `Status: ${state.statusMsg}`}
              </div>
              {/* <div className="cstformHdrMiddleContainer mobileTitleNewForm">
                eNote - {isNewForm ? "New" : `${state.noteId}`}
              </div> */}
              <div className="cstformHdrMiddleContainer mobileTitleNewForm">
                {(() => {
                  const returnedId = enumsObj?.NoteStatus.find(x => x.dValue === "Returned")?.id;
                  const calledBackId = enumsObj?.NoteStatus.find(x => x.dValue === "Called Back")?.id;
                  if (!isNewForm && (state.status === returnedId || state.status === calledBackId)) {
                    return `eNote - ${state.noteNumber}`;
                  }
                  return `eNote - ${isNewForm ? "New" : `${state.noteId}`}`;
                })()}
              </div>
              <div className="cstformHdrRightContainer mobileTitleNewForm">
                {isNewForm ? (
                  <span>
                    Date: <Clock format={"DD-MMM-YYYY hh:mm:ss A"} ticking={true} />
                  </span>
                ) : (
                  `Created: ${state.createdDate}`
                )}
              </div>
            </div>
          </div>
          <div className="container largeScreen mobileViewCSS">
            <Form
              render={() => (
                <FormElement>
                  <fieldset className={"k-form-fieldset"}>
                    <div className="errorMsg">All fields marked "*" are mandatory</div>
                    <div className="SectionHeads row">General Section</div>
                    <div className="SectionRow row">
                      <div className="col-md-6">
                        <FieldWrapper>
                          <div className="k-form-field-wrap">
                            <Label className="k-form-label">
                              Department:
                              <span className="required-asterisk">*</span>
                            </Label>

                            {userDepartment && (
                              <Field
                                component={Input}
                                name="Department"
                                value={state.Department}
                                defaultValue={userDepartment}
                                className="_departmentBorder"
                                readOnly
                              />
                            )}

                          </div>
                          <div className="errorMsg">{departmentError}</div>
                        </FieldWrapper>
                      </div>

                      <div className="col-md-6">
                        <FieldWrapper>
                          <div className="k-form-field-wrap">
                            <Label className="k-form-label">
                              Note To:<span className="required-asterisk">*</span>
                            </Label>
                            <DropDownList
                              data={noteto.map((x) => x.dValue)}
                              onChange={handleNote}
                              value={state.NoteTo}
                              name="NoteTo"
                              style={{ borderColor: noteBorderColor }}
                              valueRender={element => valueRender(element, state.NoteTo, 'NoteTo')}
                            />
                          </div>
                        </FieldWrapper>
                      </div>
                      <div className="col-md-6">
                        <FieldWrapper>
                          <div className="k-form-field-wrap">
                            <Label className="k-form-label">
                              Subject:<span className="required-asterisk">*</span>
                            </Label>
                            <TextArea
                              maxLength={max}
                              onChange={handleSubjectChange}
                              onKeyDown={(e) => {
                                // TEMP FIX: Allow space
                                if (e.key === ' ') {
                                  e.stopPropagation(); // cancel any internal block
                                }
                              }}
                              value={state.Subject}
                              rows={1}
                              style={{ borderColor: subjectBorderColor }}
                              className="mobileFieldHeight"
                            />

                            {subjectErrorMessage && (
                              <div className="inCorrectFileError">
                                {subjectErrorMessage}
                              </div>
                            )}
                            <Hint direction={"end"}>
                              {state.Subject.length} / {max}
                            </Hint>
                          </div>
                        </FieldWrapper>
                      </div>
                      <div className="col-md-6">
                        <FieldWrapper>
                          <div className="k-form-field-wrap">
                            <Label className="k-form-label">
                              Nature of Note:
                              <span className="required-asterisk">*</span>
                            </Label>
                            <DropDownList
                              name="NatureOfNote"
                              data={natureofnote.map((x) => x.dValue)}
                              value={state.NatureOfNote}
                              onChange={handleNatureOfNoteChange}
                              style={{ borderColor: natureofnoteBorderColor }}
                              valueRender={element => valueRender(element, state.NatureOfNote, 'NatureOfNote')}
                            />
                          </div>
                        </FieldWrapper>
                      </div>
                      {showNatureOfApprSanc && (
                        <div className="col-md-6">
                          <FieldWrapper>
                            <div className="k-form-field-wrap">
                              <Label className="k-form-label">
                                Nature of Approval/Sanction:
                                <span className="required-asterisk">*</span>
                              </Label>
                              <DropDownList
                                name="NatureOfApprSanc"
                                value={state.NatureOfApprSanc}
                                data={natureofapprsanc.map((x) => x.dValue)}
                                onChange={handleNatureOfApprSancChange}
                                style={{ borderColor: natureofapprsancBorderColor }}
                                valueRender={element => valueRender(element, state.NatureOfApprSanc, 'NatureOfApprSanc')}
                              />
                            </div>
                          </FieldWrapper>
                        </div>
                      )}
                      <div className="col-md-6">
                        <FieldWrapper>
                          <div className="k-form-field-wrap">
                            <Label className="k-form-label">
                              Note Type:
                              <span className="required-asterisk">*</span>
                            </Label>
                            <DropDownList
                              data={notetype.map((x) => x.dValue)}
                              value={state.NoteType}
                              onChange={handleNoteType}
                              style={{ borderColor: notetypeBorderColor }}
                              valueRender={element => valueRender(element, state.NoteType, 'NoteType')}
                            />
                          </div>
                        </FieldWrapper>
                      </div>
                      {showTypeOfFinNote && (
                        <div className="col-md-6">
                          <FieldWrapper>
                            <div className="k-form-field-wrap">
                              <Label className="k-form-label">
                                Type of Financial Note:
                                <span className="required-asterisk">*</span>
                              </Label>
                              <DropDownList
                                data={fintype.map((x) => x.dValue)}
                                value={state.TypeOfFinNote}
                                onChange={handleTypeOfFin}
                                style={{ borderColor: typeoffinBorderColor }}
                                valueRender={element => valueRender(element, state.TypeOfFinNote, 'TypeOfFinNote')}
                              />
                            </div>
                          </FieldWrapper>
                        </div>
                      )}
                      {showAmount && (
                        <div className="col-md-6">
                          <FieldWrapper>
                            <div className="k-form-field-wrap">
                              <Label className="k-form-label">
                                Amount:<span className="required-asterisk">*</span>
                              </Label>
                              <Input
                                value={amount}
                                onChange={handleAmount}
                                defaultValue={amount}
                                style={{ borderColor: amountBorderColor }}
                              />
                            </div>
                          </FieldWrapper>
                        </div>
                      )}
                      <div className="col-md-6">
                        <FieldWrapper>
                          <div className="k-form-field-wrap">
                            <Label className="k-form-label">
                              Search Text:<span className="required-asterisk">*</span>
                            </Label>
                            <TextArea
                              maxLength={max}
                              value={state.Search}
                              onChange={handleSearchChange}
                              rows={1}
                              style={{ borderColor: searchBorderColor }}
                              className="mobileFieldHeight"
                            />
                            {searchErrorMessage && (
                              <div className="inCorrectFileError">{searchErrorMessage}</div>
                            )}
                            <Hint direction={"end"}>{state.Search.length} / {max}</Hint>
                          </div>
                        </FieldWrapper>
                      </div>
                      {showPurpose && (
                        <div className="col-md-6">
                          <FieldWrapper>
                            <div className="k-form-field-wrap">
                              <Label className="k-form-label">
                                Purpose:
                                <span className="required-asterisk">*</span>
                              </Label>
                              <Input
                                value={purpose}
                                maxLength={maxPurpose}
                                name="Purpose"
                                onChange={handlePurpose}
                                style={{ borderColor: purposeBorderColor }}
                                className="mobileFieldHeight"
                              />
                              {purposeErrorMessage && (
                                <div className="inCorrectFileError">
                                  {purposeErrorMessage}
                                </div>
                              )}
                              <Hint direction={"end"}>
                                {(purpose || "").length} / {maxPurpose}
                              </Hint>
                            </div>
                          </FieldWrapper>
                        </div>
                      )}
                      {showPurpose1 && (
                        <div className="col-md-6">
                          <FieldWrapper>
                            <div className="k-form-field-wrap">
                              <Label className="k-form-label">
                                Purpose:
                                <span className="required-asterisk">*</span>
                              </Label>
                              <DropDownList
                                data={drpOptPurposeApproval.map((x) => x.dValue)}
                                value={state.PurposeApproval}
                                onChange={handlePurposeApproval}
                                style={{ borderColor: purpose1BorderColor }}
                                valueRender={element => valueRender(element, state.PurposeApproval, 'PurposeApproval')}
                              />
                            </div>
                          </FieldWrapper>
                        </div>
                      )}
                      {showPurpose2 && (
                        <div className="col-md-6">
                          <FieldWrapper>
                            <div className="k-form-field-wrap">
                              <Label className="k-form-label">
                                Purpose:
                                <span className="required-asterisk">*</span>
                              </Label>
                              <DropDownList
                                data={drpOptPurposeInformation.map(
                                  (x) => x.dValue
                                )}
                                value={state.PurposeInformation}
                                onChange={handlePurposeInformation}
                                style={{ borderColor: purpose2BorderColor }}
                                valueRender={element => valueRender(element, state.PurposeInformation, 'PurposeInformation')}
                              />
                            </div>
                          </FieldWrapper>
                        </div>
                      )}
                      {(showPurpose1 && purpose === "Others") && (
                        <div className="col-md-6">
                          <FieldWrapper>
                            <div className="k-form-field-wrap">
                              <Label className="k-form-label">
                                Others:<span className="required-asterisk">*</span>
                              </Label>
                              <Input
                                value={purposeOthers}
                                name="PurposeOthers"
                                onChange={handlePurposeInfoOthers}
                                style={{ borderColor: purposeOthersBorderColor }}
                                className="mobileFieldHeight"
                              />
                              <Hint direction={"end"}>
                                {(purposeOthers || "").length} / {maxPurpose}
                              </Hint>
                            </div>
                          </FieldWrapper>
                        </div>
                      )}
                      {showConfidential && (
                        <div className="col-md-6">
                          <FieldWrapper>
                            <div className="k-form-field-wrap">
                              <Label className="k-form-label">
                                Confidential Note:
                              </Label>
                              <Switch
                                checked={isConfidential}
                                onChange={(e) => setIsConfidential(e.value)}
                                onLabel="Yes"
                                offLabel="No"
                              />
                            </div>
                          </FieldWrapper>
                        </div>
                      )}
                    </div>

                    <div className="SectionHeads row">Approver Details</div>
                    <div className="SectionRow row">
                      <div className="_approverDetailsDiv">
                        <MultiColumnComboBox
                          data={orgEmployees}
                          filterable={true}
                          // mobile view design for dropdown
                          columns={isMobile ? mobileColumns : orgUsersPplPickerColumnMapping}
                          value={
                            selectedReveiwer === null
                              ? null
                              : selectedReveiwer.displayName
                          }
                          onFilterChange={pplFilterMultiColumn}
                          onChange={handleComboChangeReviewer}
                          onOpen={handleOpenReveiwer}
                          style={{ width: isMobile ? "100%" : "300px", marginRight: "5px", }} // Adjust width
                          placeholder="Add Reviewer..."

                        />
                        <Button onClick={handleAddRow}>
                          <span className="k-icon k-font-icon k-i-plus cursor allIconsforPrimary-btn"></span>
                          Add
                        </Button>
                      </div>
                      {/* Added this message based on change happend due to loading impact users will be filtered on search box enter - 22/03 */}
                      <div className="cstUserSearchMsg">(Please enter minimum 4 characters to search)</div>
                      <div className="cstDisableGridScroll">
                        {/* added scroll class 24/04 */}
                        <TableDraggableRows
                          srCol="Employee Id"
                          designationCol="Designation"
                          reviewerCol="Reviewer"
                          data={notereviewerData}
                          onDelate={deletereviewer}
                          onOrderChange={orderChange} />
                      </div>
                      <div className="_reviewerDetailsDiv">
                      </div>
                      <div className="_reviewerDetailsCol">
                        <MultiColumnComboBox
                          data={orgEmployees}
                          filterable={true}
                          // mobile view design for dropdown
                          columns={isMobile ? mobileColumns : orgUsersPplPickerColumnMapping}
                          style={{ width: isMobile ? "100%" : "300px", marginRight: "5px", }} // Adjust width
                          value={
                            combovalueApprover === null
                              ? null
                              : combovalueApprover.displayName
                          }
                          onFilterChange={pplFilterMultiColumn}
                          onChange={handleComboChangeApprover}
                          onOpen={handleOpenApprover}
                          placeholder="Add Approver..."
                        />
                        <Button onClick={handleAddRowApprover}>
                          <span className="k-icon k-font-icon k-i-plus cursor allIconsforPrimary-btn"></span>
                          Add</Button>
                      </div>
                      {/* Added this message based on change happend due to loading impact users will be filtered on search box enter - 22/03 */}
                      <div className="cstUserSearchMsg">(Please enter minimum 4 characters to search)</div>
                      <div className="cstDisableGridScroll">
                        {/* added scroll class 24/04 */}
                        <TableDraggableRows
                          srCol="Employee Id"
                          designationCol="Designation"
                          reviewerCol="Approver"
                          data={noteapproverData}
                          onDelate={deletereviewer}
                          onOrderChange={orderChange} />
                      </div>
                      <div className="_reviewerDetailsDiv">
                      </div>
                    </div>

                    {(!isNewForm && state.status === enumsObj?.NoteStatus.find(x => x.dValue === "Returned").id) &&
                      (<div>
                        <div className="SectionHeads eNoteAtrFormHdr row">
                          Comments
                        </div>
                        <div className="SectionRow row">
                          <div className="table-responsive">
                            <table className="SectionRow cstATRFormTbl tableStyle">
                              <tr>
                                <th className="approvalform-tableCol-width-1">Page#</th>
                                <th className="approvalform-tableCol-width-1">Doc Reference</th>
                                <th className="approvalform-tableCol-width-5">Comments</th>
                                <th className="approvalform-tableCol-width-3">Comment By</th>
                              </tr>
                              {noteComments?.map((comment) => (
                                <tr key={comment.noteApproverCommentID}>
                                  <td className="approvalform-tableCol-width-1">{comment.pageNumber}</td>
                                  <td className="approvalform-tableCol-width-1">{comment.docReferrence}</td>
                                  <td className="approvalform-tableCol-width-5">{comment.comments}</td>
                                  <td className="approvalform-tableCol-width-3">{comment.approverEmailName}</td>
                                </tr>
                              )
                              )}
                            </table>
                          </div>
                        </div>
                      </div>
                      )}

                    <div className="SectionHeads row">File Attachments</div>
                    <div className="SectionRow row">
                      <div className="col-md-6">
                        {" "}
                        <Label className="k-form-label">
                          Note PDF<span className="required-asterisk">*</span>
                        </Label>
                        <div className="Attachemntfileinfo-ind">
                          <input
                            type="file"
                            id="PDFDocfile"
                            key={fileKey} // Unique key to force re-render
                            // onChange={(event) => convertPDFToBase64(event, 'PDF')} // PDF as file category
                            onChange={(event) => convertPDFToBase64(event.target.files[0], 'PDF')}
                            // onChange={convertPDFToBase64}
                            className="_ConvertPDFToBase64"
                          />
                        </div>
                        {wordandPdfWarring?.PDFInfo.fileName === null ||
                          wordandPdfWarring?.PDFInfo.fileName === "" ? null : (
                          <div className="Attachemntfileinfo-ind">
                            <span className="attachemntIconInfoConationer">
                              <span className="AttchemntIconWraper">
                                <SvgIcon
                                  icon={getFileIcon(
                                    wordandPdfWarring?.PDFInfo.fileName
                                  )}
                                  size="xxlarge"
                                />
                                <span className="attachemnrt-warningifoConatiner">
                                  <div className="attachemnrt-warningifo-fileinfo">
                                    {wordandPdfWarring?.PDFInfo.fileName}
                                  </div>
                                  <span
                                    className="inCorrectFileError"
                                  >
                                    {wordandPdfWarring?.PDFInfo.warningMsg}
                                  </span>
                                </span>
                              </span>
                              <span
                                className="AttchemntIconWraperCancel"
                                onClick={() =>
                                  onRemoveAttachmentWarning("PDFInfo")
                                }
                              >
                                X
                              </span>
                            </span>
                          </div>
                        )}
                        <div className="_fileFormatHintMsg">
                          Allowed only one PDF.Upto 25MB max.
                        </div>
                      </div>
                      {noteapproverData?.length > 0 && isSecretaryExist && (
                        <div className="col-md-6">
                          <Label className="k-form-label">
                            Word Document
                            <span className="required-asterisk">*</span>
                          </Label>

                          <div className="Attachemntfileinfo-ind">
                            {" "}
                            <input
                              type="file"
                              id="WordDocfile"
                              key={fileKey} // Unique key to force re-render
                              // onChange={(event) => convertBase64Word(event.target.files[0], 'WORD')}
                              // onChange={(event) => convertBase64Word(event, 'WORD')} // WORD as file category
                              onChange={convertBase64Word}
                              className="_ConvertPDFToBase64"
                            />
                          </div>
                          {wordandPdfWarring?.wordInfo.fileName === null ||
                            wordandPdfWarring?.wordInfo.fileName === "" ? null : (
                            <div className="Attachemntfileinfo-ind">
                              <span className="attachemntIconInfoConationer">
                                <span className="AttchemntIconWraper">
                                  <SvgIcon
                                    icon={getFileIcon(
                                      wordandPdfWarring?.wordInfo.fileName
                                    )}
                                    size="xxlarge"
                                  />
                                  <span className="attachemnrt-warningifoConatiner">
                                    <div className="attachemnrt-warningifo-fileinfo">
                                      {wordandPdfWarring?.wordInfo.fileName}
                                    </div>
                                    <span
                                      className="inCorrectFileError"
                                    >
                                      {wordandPdfWarring?.wordInfo.warningMsg}
                                    </span>
                                  </span>
                                </span>
                                <span
                                  className="AttchemntIconWraperCancel"
                                  onClick={() =>
                                    onRemoveAttachmentWarning("wordInfo")
                                  }
                                >
                                  X
                                </span>
                              </span>
                            </div>
                          )}

                          <div className="_fileFormatHintMsg">
                            Allowed only one Word.Upto 10MB max.
                          </div>
                        </div>
                      )}
                      <div className="col-md-6">
                        <Label className="k-form-label">
                          Supporting Documents
                        </Label>
                        {/* multidoc controller */}
                        <div className="Attachemntfileinfo-ind">
                          <input
                            type="file"
                            id="multiDoc"
                            key={fileKey} // Unique key to force re-render
                            // onChange={(event) => multipleDocUpload(event.target.files[0], 'SupportingDocument')}
                            // onChange={(event) => multipleDocUpload(event, 'SupportingDocument')} // Supporting Document as file category
                            onChange={multipleDocUpload}
                            className="_ConvertPDFToBase64"
                          />
                          <div>
                            <span className="inCorrectFileError">
                              {supportingDocError}
                            </span>
                          </div>
                        </div>
                        {filesInfo?.map((obj, ind) => (
                          <div className="Attachemntfileinfo-ind" key={ind}>
                            <span className="attachemntIconInfoConationer">
                              <span className="AttchemntIconWraper">
                                {/* Assuming you have different icons for different file types */}
                                <SvgIcon
                                  icon={getFileIcon(
                                    obj.supportingDocumentFileName
                                  )}
                                  size="xxlarge"
                                />
                                <span className="attachemnrt-warningifoConatiner">
                                  <div className="attachemnrt-warningifo-fileinfo">
                                    {obj.supportingDocumentFileName}
                                  </div>
                                  <span className="inCorrectFileError">
                                    {
                                      supportDocWarning[
                                        obj.supportingDocumentFileName
                                      ]?.warningMsg
                                    }
                                  </span>
                                </span>
                              </span>
                              <span
                                className="AttchemntIconWraperCancel"
                                onClick={() => onRemoveMultiAttachment(ind)}
                              >
                                X
                              </span>
                            </span>
                          </div>
                        ))}

                        <div className="_fileFormatHintMsg">
                          Allowed Formats (pdf,doc,docx,xlsx only) Upto 25MB max.
                        </div>
                      </div>
                    </div>
                  </fieldset>
                  <div className="FormButtonsContainer k-form-buttons">
                    {(isNewForm ||
                      (!isNewForm &&
                        state.status === 1 &&
                        state.createdBy === accounts[0].username)) && (
                        <Button
                          type={"button"}
                          onClick={handleSave}
                          className="FormButtons k-button k-button-md k-rounded-md k-button-solid k-button-solid-base"
                        >
                          <span className=" k-icon-xs k-icon k-font-icon k-i-save cursor allIconsforPrimary-btn"></span>
                          <save>
                            Save as Draft
                          </save>
                        </Button>
                      )}
                    {(isNewForm ||
                      (!isNewForm &&
                        (state.status === 1 ||
                          state.status === 4 ||
                          state.status === 8) &&
                        state.createdBy === accounts[0].username)) && (
                        <Button
                          onClick={handleOpenDialog}
                          type={"submit"}
                          className="FormButtons k-button k-button-md k-rounded-md k-button-solid k-button-solid-base"
                        >
                          <span className="k-icon k-font-icon k-i-launch cursor allIconsforPrimary-btn"></span>
                          Submit
                        </Button>
                      )}
                    {!isNewForm &&
                      state.status === 4 &&
                      state.createdBy === accounts[0].username && (
                        <span className="eNote-ApprovalButton">
                          <Button
                            onClick={() => setVisibleCancelCfrm(true)}
                            className="formBtnColor"
                          >
                            <span className="k-icon-sm k-icon k-font-icon k-i-cancel cursor allIconsforPrimary-btn"></span>
                            Cancel Note
                          </Button>
                        </span>
                      )}
                    <Link to="/datagridpage">
                      <Button
                        type={"submit"}
                        className="k-button k-button-md k-rounded-md k-button-solid k-button-solid-base"
                      >
                        <span className="k-icon-sm k-icon k-font-icon  k-i-x-circle cursor allIconsforPrimary-btn"></span>
                        Exit
                      </Button>
                    </Link>
                  </div>
                  {isLoading && <PageLoader />}
                  {visiblesave && (
                    <Dialog
                      title={<CustomDialogTitleBar />}
                      onClose={handleSaveClose}
                    >
                      <p className="dialogcontent_">
                        {successMessage}
                      </p>
                      <DialogActionsBar>
                        <Button
                          className="notifyDailogOkBtn k-button k-button-md k-rounded-md k-button-solid k-button-solid-base"
                          onClick={() => { setTab("My Pending Notes"); navigate(redirectTo); }}
                        >
                          <span className="k-icon k-font-icon  k-i-redo cursor allIconsforPrimary-btn"></span>
                          Ok
                        </Button>
                      </DialogActionsBar>
                    </Dialog>
                  )}
                  {visible && (
                    <Dialog
                      title={<CustomConfirmDialogTitleBar />}
                      onClose={handleCloseDialog}
                      className="dialogcontent_Refer"
                    >
                      <p className="dialogcontent_">
                        Are you sure you want to submit this request?
                      </p>
                      <p className="dialogcontent_">
                        Please check the details filled along with attachment and
                        click on Confirm button to submit request.
                      </p>
                      <DialogActionsBar>
                        <Button
                          className="k-button k-button-md k-rounded-md k-button-solid k-button-solid-base formBtnColor"
                          onClick={handleSubmit}
                        >
                          {/* {isLoading ? 'Loading...' : 'Confirm'} */}
                          <span className="k-icon k-font-icon  k-i-checkmark-circle cursor allIconsforPrimary-btn"></span>
                          Confirm
                        </Button>
                        <Button
                          className="k-button k-button-md k-rounded-md k-button-solid k-button-solid-base"
                          onClick={handleCloseDialog}
                        >
                          <span className="k-icon k-font-icon  k-i-close-circle cursor allIconsforPrimary-btn"></span>
                          Cancel
                        </Button>
                      </DialogActionsBar>
                    </Dialog>
                  )}
                  {visibleCancelCfrm && (
                    <Dialog
                      title={<CustomConfirmDialogTitleBar />}
                      onClose={() => setVisibleCancelCfrm(false)}
                      className="dialogcontent_Refer"
                    >
                      <p className="dialogcontent_">
                        Are you sure you want to cancel this request?
                      </p>
                      <p className="dialogcontent_">
                        Please click on Confirm button to cancel request.
                      </p>
                      <DialogActionsBar>
                        <Button
                          className="k-button k-button-md k-rounded-md k-button-solid k-button-solid-base formBtnColor"
                          onClick={onCancelNote}
                        >
                          <span className="k-icon k-font-icon  k-i-checkmark-circle cursor allIconsforPrimary-btn"></span>
                          Confirm
                        </Button>
                        <Button
                          className="k-button k-button-md k-rounded-md k-button-solid k-button-solid-base"
                          onClick={() => setVisibleCancelCfrm(false)}
                        >
                          <span className="k-icon k-font-icon  k-i-close-circle cursor allIconsforPrimary-btn"></span>
                          Cancel
                        </Button>
                      </DialogActionsBar>
                    </Dialog>
                  )}
                  {alertvisible && (
                    <Dialog
                      title={<CustomDialogTitleBar />}
                      onClose={() => setAlertVisible(false)}
                    >
                      <p className="dialogcontent_">
                        The request for eNote has been submitted successfully.
                      </p>
                      <DialogActionsBar>
                        <Button
                          className="notifyDailogOkBtn k-button k-button-md k-rounded-md k-button-solid k-button-solid-base"
                          onClick={() => { setTab("My Pending Notes"); navigate(redirectTo); }}
                        >
                          <span className="k-icon k-font-icon  k-i-redo cursor allIconsforPrimary-btn"></span>
                          Ok
                        </Button>
                        {/* </Link> */}
                      </DialogActionsBar>
                    </Dialog>
                  )}

                  {validationErrors && (
                    <Dialog
                      title={<CustomDialogTitleBar />}
                      onClose={handleClosevalidationDialog}
                    >
                      <p className="cstDailogValidmsg">
                        {/* Bug fix - order of validation - 05/04 - RK */}
                        <div className="_cstDialogDiv">
                          <p className="_cstDialogDivp">Please fill up all the mandatory fields</p>
                          <ul className="_cstDialogDivUl">
                            {errorMessages.map((message, index) => (
                              <li className="_cstDialogDivUlli" key={index}>{message}</li>
                            ))}
                          </ul>
                          <p>
                            <strong>Note:</strong> Invalid files are not allowed
                          </p>
                        </div>
                      </p>
                      <DialogActionsBar>
                        <Button
                          onClick={handleClosevalidationDialog}
                          className="notifyDailogOkBtn k-button k-button-md k-rounded-md k-button-solid k-button-solid-base"
                        >
                          <span className="k-icon k-font-icon  k-i-redo cursor allIconsforPrimary-btn"></span>
                          Ok
                        </Button>
                      </DialogActionsBar>
                    </Dialog>
                  )}

                  {showNotification && (
                    <Dialog
                      title={<CustomDialogTitleBar />}
                      onClose={() => setShowNotification(false)}
                    >
                      <p className="dialogcontent_">
                        {notificationMsg}
                      </p>
                      <DialogActionsBar>
                        <Button
                          onClick={() => {
                            setShowNotification(false);
                            setSelectedReviewer(null); // Clear selectedReviewer on Ok click
                          }}
                          className="notifyDailogOkBtn k-button k-button-md k-rounded-md k-button-solid k-button-solid-base"
                        >
                          <span className="k-icon k-font-icon  k-i-redo cursor allIconsforPrimary-btn"></span>
                          Ok
                        </Button>
                      </DialogActionsBar>
                    </Dialog>
                  )}
                </FormElement>
              )}
            />
          </div>
        </div>
        <Footer /></>}
    </div>
  );
};