

import React, { useState, useEffect, useRef, useCallback, useLayoutEffect, useMemo } from "react";
// import Kendo Components
import {
  ExpansionPanel,
  ExpansionPanelContent,
} from "@progress/kendo-react-layout";
import { ComboBox } from "@progress/kendo-react-dropdowns";
import {
  filePdfIcon,
  fileWordIcon,
  fileImageIcon,
  fileTxtIcon,
  fileDataIcon,
  fileIcon,
  infoCircleIcon,
} from "@progress/kendo-svg-icons";
import { MultiColumnComboBox } from "@progress/kendo-react-dropdowns";
import { Reveal } from "@progress/kendo-react-animation";
import { Button } from "@progress/kendo-react-buttons";
import { Dialog, DialogActionsBar } from "@progress/kendo-react-dialogs";
// ========== PDF ANNOTATION CHANGE START ==========
// NEW - Importing PdfAnnotator component for PDF highlighting and commenting functionality
import PdfAnnotator from '../components/PdfAnnotator.jsx';
// ========== PDF ANNOTATION CHANGE END ==========

import { RadioButton } from "@progress/kendo-react-inputs";
import { Hint } from "@progress/kendo-react-labels";
// import kendo icons and images
import { SvgIcon } from "@progress/kendo-react-common";
import view from "../assets/view.png";
import hide from "../assets/hide.png";
// import external components
import { useMsal, useAccount } from "@azure/msal-react";
import { useParams } from "react-router-dom";
import { useNavigate } from "react-router-dom";
import DateObject from "react-date-object";
import CryptoJS from "crypto-js";
import * as pdfjsLib from "pdfjs-dist";
import ReactPDFViewerApp from '../components/pdfViewerDEV'
import pdfjsWorker from "pdfjs-dist/build/pdf.worker.entry";
// import { v4 as uuidv4 } from 'uuid';
// import internal componnets
import Navbar from "../components/navbar";
import { Sidebar } from "../components/sidebar";
import Footer from "../components/footer";
import { getAccessToken, useTabContext } from "../App";
import { loginRequest, API_COMMON_HEADERS, API_BASE_URL, API_ENDPOINTS } from "../config";
import { enoteviewformitems } from "./datagridlistitems";
import Unauthorized from "../components/Unauthorized";
import { MenuNavContainer } from "../components/menu";
import { LoadingOverlay } from "../components/progressBar";
import ReactPDFViewerGist from "../components/gistPDFViewer";
// Import css styles
import "../styles/forms.css";
import "../styles/passcode.css";
//import "pdfjs-dist/web/pdf_viewer.css";
import "../styles/pdf_viewer.css";
import "../styles/responsiveDesign.css";
import "../styles/homepage.css";
import { Grid, GridColumn as Column } from "@progress/kendo-react-grid";


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
  {
    field: "srNo",
    header: "Employee Id",
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
  {
    field: "srNo",
    header: "Employee Id",
    width: "70px",
  },
];
export const ENoteViewForm = () => {
  const id = useParams();
  const navigate = useNavigate();
  const { setTab } = useTabContext();
  const { setPasscodeNavigate } = useTabContext();
  const { accounts, instance } = useMsal(); //for mail username
  const [expanded, setExpanded] = useState(["General Section"]);
  const [noteData, setNoteData] = useState(null);
  const [redirect] = useState("/datagridpage");
  const account = useAccount(accounts[0] || {});
  // general Commenta
  const [generalcmtInfoobj, SetGeneralCmtInfoObj] = useState({
    DocRef: "",
    Comments: "",
    PageNo: "",
    iEdit: false,
  });
  // add refer comments
  const [referCommentsObj, setReferCommentsObj] = useState({
    DocRef: "",
    Comments: "",
    PageNo: "",
    isEdit: false,
  });
  // dialog box for refer and  change Approver
  const [dialogForRefereeAndChangeApprover, setDialogForRefereeAndChangeApprover] = useState(false);
  const [generalcmtforAllCmt, SetGeneralcmtforAllCmt] = useState([]);  // General commnets info array
  const [aTRCreaters, SetATRCreaters] = useState([]);  // ATR Creaters details
  const [assigneeUsers, setAssigneeUsers] = useState([]); // Assigness list
  const [internalUsers, setInternalUsers] = useState([]);
  const [assigneerUserDetails, setAssigneerUserDetails] = useState([]); // dupliacte for atr assigness list
  const [selectedAssignee, setSelectedAssignee] = useState("");// combobox value for selected assignee
  const [selectedAssigneeUsersinfo, setSelectedAssigneeUsersinfo] = useState([]); // Selected assignees details
  const [supportDocfilesInfo, setSupportDocfilesInfo] = useState([]); //  supportDoc files info  //
  //word doc and gist doc details
  const [wordandPdfInfo, SetWordPDFInfowarring] = useState({
    wordInfo: {
      fileExtension: "",
      fileName: "",
      warningMsg: "",
      filePath: "",
      isValid: false,
      isDownloadble: false,
      base64: "",
    },
  });
  const [fileWordWarning, SetfilesWordWarning] = useState(""); // word invaild file error message info //
  const [supportDocWarning, SetSupportDocWarning] = useState({}); // support doc warning msg
  // noteWord
  const [noteWordDoc, setNoteWordDoc] = useState({
    noteWordDocInfo: {
      fileName: "",
      warningMsg: "",
      filePath: "",
      isValid: false,
    },
  });
  //expend view json;
  const [expendJson, SetExpendJson] = useState(enoteviewformitems);
  const [isSecretary, setIsSecretary] = useState(false);
  const [isCurrentATRCreator, setIsCurrentATRCreator] = useState(false);
  const [isreferredUser, SetIsreferredUser] = useState(false);


  // ========== PDF ANNOTATION CHANGE START ==========
  // NEW - State to store PDF annotation comments extracted from PDF viewer
  const [pdfAnnotationComments, setPdfAnnotationComments] = useState([]);
  const [includePdfAnnotations, setIncludePdfAnnotations] = useState(true);
  const pdfAnnotatorRef = useRef(null);
  const [pdfInitialAnnotations, setPdfInitialAnnotations] = useState([]);
  // ========== PDF ANNOTATION CHANGE END ==========




  // ========== PDF ANNOTATION CHANGE START ==========
  // NEW - Generates unique localStorage key for storing draft annotations per note and user
  // Format: "enote_pdf_annotations_{noteId}_{userEmail}"
  const getPdfDraftStorageKey = useCallback((noteId, email) => {
    return `enote_pdf_annotations_${noteId}_${(email || "").toLowerCase()}`;
  }, []);

  const safeJsonParse = (value, fallback = null) => {
    try {
      return JSON.parse(value);
    } catch {
      return fallback;
    }
  };

  const loadCurrentUserDraftAnnotations = useCallback(() => {
    const currentEmail = accounts?.[0]?.username;
    const currentNoteId = noteData?.noteId || id?.id;

    if (!currentEmail || !currentNoteId) return [];

    const saved = localStorage.getItem(
      getPdfDraftStorageKey(currentNoteId, currentEmail)
    );

    const parsed = safeJsonParse(saved, []);
    return Array.isArray(parsed) ? parsed : [];
  }, [accounts, noteData, id, getPdfDraftStorageKey]);

  const saveCurrentUserDraftAnnotations = useCallback((annotations) => {
    const currentEmail = accounts?.[0]?.username;
    const currentNoteId = noteData?.noteId || id?.id;

    if (!currentEmail || !currentNoteId) return;

    const onlyCurrentUserAnnotations = (Array.isArray(annotations) ? annotations : []).filter(
      (item) =>
        (item?.user || "").toLowerCase() ===
        (currentEmail || "").toLowerCase()
    );

    localStorage.setItem(
      getPdfDraftStorageKey(currentNoteId, currentEmail),
      JSON.stringify(onlyCurrentUserAnnotations)
    );
  }, [accounts, noteData, id, getPdfDraftStorageKey]);

  const clearCurrentUserDraftAnnotations = useCallback(() => {
    const currentEmail = accounts?.[0]?.username;
    const currentNoteId = noteData?.noteId || id?.id;

    if (!currentEmail || !currentNoteId) return;

    localStorage.removeItem(getPdfDraftStorageKey(currentNoteId, currentEmail));
  }, [accounts, noteData, id, getPdfDraftStorageKey]);


  const shouldHideAnnotations = useMemo(() => {
    const currentStatus = Number(noteData?.status || noteData?.noteStatus || 0);

    // Final statuses
    return [5, 6, 7].includes(currentStatus); // Approved, Rejected, Cancelled
  }, [noteData]);













  // ========== PDF ANNOTATION CHANGE END ==========

  // ========== PDF ANNOTATION CHANGE START ==========
  // NEW - Extracts PDF highlights and comments from API response data
  const extractPdfAnnotationsFromAPI = useCallback(() => {
    const allAnnotationDetails = Array.isArray(noteData?.annotationDetailsDTO)
      ? noteData.annotationDetailsDTO
      : [];

    const allComments = Array.isArray(noteData?.noteApproverCommentsDTO)
      ? noteData.noteApproverCommentsDTO
      : [];

    const annotationMap = new Map();

    const commentByAnnotationId = new Map();

    allComments.forEach((comment) => {
      const raw = comment?.comments || "";
      const annotationId = comment?.annotationId || comment?.AnnotationId;

      if (!annotationId) return;

      if (raw.startsWith("[PDF Comment]:")) {
        commentByAnnotationId.set(Number(annotationId), comment);
      }
    });

    const groupMap = new Map();

    allAnnotationDetails.forEach((detail) => {
      const annotationId = detail.annotationId || detail.AnnotationId;
      if (!annotationId) return;

      const groupId =
        detail.highlightGroupId ||
        detail.HighlightGroupId ||
        `single-${annotationId}`;

      if (!groupMap.has(groupId)) {
        groupMap.set(groupId, []);
      }

      groupMap.get(groupId).push(detail);
    });

    groupMap.forEach((relatedDetails, groupId) => {
      relatedDetails.sort((a, b) => {
        const pageA = Number(a.pageNo || a.PageNo || 0);
        const pageB = Number(b.pageNo || b.PageNo || 0);
        if (pageA !== pageB) return pageA - pageB;

        const yA = Number(a.yposition || a.Yposition || 0);
        const yB = Number(b.yposition || b.Yposition || 0);
        if (yA !== yB) return yA - yB;

        const xA = Number(a.xposition || a.Xposition || 0);
        const xB = Number(b.xposition || b.Xposition || 0);
        return xA - xB;
      });

      const mainAnnotation = relatedDetails[0];
      const annotationId = mainAnnotation.annotationId || mainAnnotation.AnnotationId;

      const matchingComment = commentByAnnotationId.get(Number(annotationId));

      const page = Number(mainAnnotation.pageNo || mainAnnotation.PageNo || 1);
      const createdBy = mainAnnotation.createdBy || mainAnnotation.CreatedBy || "User";

      const timestamp =
        mainAnnotation.createdDate ||
        mainAnnotation.CreatedDate ||
        mainAnnotation.modifiedDate ||
        mainAnnotation.ModifiedDate ||
        new Date().toISOString();

      const selectedText =
        mainAnnotation.selectedText ||
        mainAnnotation.SelectedText ||
        "";

      const normalizedRects = relatedDetails.map((detail) => ({
        x: Number(detail.xposition || detail.Xposition || 0),
        y: Number(detail.yposition || detail.Yposition || 0),
        width: Number(detail.width || detail.Width || 0),
        height: Number(detail.height || detail.Height || 0),
      }));

      if (matchingComment) {
        const commentText = (matchingComment.comments || "")
          .replace("[PDF Comment]:", "")
          .trim();

        annotationMap.set(groupId, {
          id: groupId,
          type: "comment",
          page,
          text: commentText,
          selectedText,

          lineNumber:
            matchingComment.docReferrence ||
            matchingComment.DocReferrence ||
            "NA",

          lineNumberRange:
            matchingComment.docReferrence ||
            matchingComment.DocReferrence ||
            "NA",

          user:
            matchingComment.createdBy ||
            matchingComment.CreatedBy ||
            matchingComment.approverEmail ||
            createdBy,

          timestamp:
            matchingComment.createdDate ||
            matchingComment.CreatedDate ||
            timestamp,

          annotationId,
          highlightGroupId:
            mainAnnotation.highlightGroupId ||
            mainAnnotation.HighlightGroupId ||
            null,

          rects: [],
          normalizedRects,

          xposition: mainAnnotation.xposition || mainAnnotation.Xposition || 0,
          yposition: mainAnnotation.yposition || mainAnnotation.Yposition || 0,
          width: mainAnnotation.width || mainAnnotation.Width || 0,
          height: mainAnnotation.height || mainAnnotation.Height || 0,

          canEdit:
            (
              matchingComment.createdBy ||
              matchingComment.CreatedBy ||
              createdBy ||
              ""
            ).toLowerCase() ===
            (accounts?.[0]?.username || "").toLowerCase(),
        });
      } else {
        const fillColor = (
          mainAnnotation.fillColor ||
          mainAnnotation.FillColor ||
          ""
        ).toLowerCase();
        if (fillColor !== "yellow") {
          return;
        }
        annotationMap.set(groupId, {
          id: groupId,
          type: "highlight",
          page,
          selectedText,

          lineNumber: "NA",
          lineNumberRange: "NA",

          user: createdBy,
          timestamp,

          annotationId,
          highlightGroupId:
            mainAnnotation.highlightGroupId ||
            mainAnnotation.HighlightGroupId ||
            null,

          rects: [],
          normalizedRects,

          xposition: mainAnnotation.xposition || mainAnnotation.Xposition || 0,
          yposition: mainAnnotation.yposition || mainAnnotation.Yposition || 0,
          width: mainAnnotation.width || mainAnnotation.Width || 0,
          height: mainAnnotation.height || mainAnnotation.Height || 0,

          canEdit:
            (createdBy || "").toLowerCase() ===
            (accounts?.[0]?.username || "").toLowerCase(),
        });
      }
    });

    return Array.from(annotationMap.values());
  }, [noteData, accounts]);

  // ========== PDF ANNOTATION CHANGE END ==========



  const getLineNumberFromComment = useCallback((comment) => {
    return (
      comment?.docReferrence ||
      comment?.DocReferrence ||
      "NA"
    );
  }, []);


  const getDisplayCommentText = useCallback((comment) => {
    const raw = comment?.comments || "";

    if (raw.startsWith("[PDF Highlight]:")) {
      return null;
    }

    if (!raw.startsWith("[PDF Comment]:")) {
      return raw;
    }

    const userComment = raw.replace("[PDF Comment]:", "").trim();
    const pageNo = comment?.pageNumber || comment?.PageNumber || "NA";
    const lineNo = getLineNumberFromComment(comment);

    return `PageNo - ${pageNo} , Line No - ${lineNo} , UserComment - ${userComment}`;
  }, [getLineNumberFromComment]);


  // ========== PDF ANNOTATION CHANGE START ==========
  // NEW - Merges annotations from API (saved) with user's draft annotations (unsaved)
  // Shows both saved and draft annotations together in PDF viewer
  const buildMergedPdfAnnotations = useCallback(() => {
    if (shouldHideAnnotations) {
      return [];
    }

    const apiAnnotations = extractPdfAnnotationsFromAPI();
    const draftAnnotations = loadCurrentUserDraftAnnotations();

    const merged = new Map();

    apiAnnotations.forEach((ann) => {
      merged.set(ann.id, {
        ...ann,
        canEdit:
          (ann?.user || "").toLowerCase() ===
          (accounts?.[0]?.username || "").toLowerCase(),
      });
    });

    draftAnnotations.forEach((ann) => {
      merged.set(ann.id, {
        ...ann,
        canEdit:
          (ann?.user || "").toLowerCase() ===
          (accounts?.[0]?.username || "").toLowerCase(),
      });
    });

    return Array.from(merged.values());
  }, [
    extractPdfAnnotationsFromAPI,
    loadCurrentUserDraftAnnotations,
    accounts,
    shouldHideAnnotations,
  ]);
  // ========== PDF ANNOTATION CHANGE END ==========

  // ========== PDF ANNOTATION CHANGE START ==========
  // NEW - Effect to load annotations when noteData is available
  // Also clears drafts if note is in final status (Approved/Rejected/Cancelled)
  useEffect(() => {
    if (!noteData) return;

    if (shouldHideAnnotations) {
      clearCurrentUserDraftAnnotations();
      setPdfInitialAnnotations([]);
      setPdfAnnotationComments([]);
      return;
    }

    setPdfInitialAnnotations(buildMergedPdfAnnotations());
  }, [
    noteData,
    buildMergedPdfAnnotations,
    shouldHideAnnotations,
    clearCurrentUserDraftAnnotations,
  ]);
  // ========== PDF ANNOTATION CHANGE END ==========

  //priti shelke - 08-04-2026


  const [actionBtn, setActionBtn] = useState("");
  const [dialogOpen, setdialogOpen] = useState(false);  // dailog for confiramation
  const [dialogSuccess, setDialogSuccess] = useState(false); // dailog for Success
  const [dialogFailure, setDialogFailure] = useState(false); // dailog for Failure
  const [getNoteSecretaryUserInfo, setGetNoteSecretaryUserInfo] = useState([]);  // getNoteSecretary
  const [orgEmployees, setOrgEmployees] = useState([]);
  const [selectedrefer, SetSelectedRefer] = useState(null); // selected refer User info
  const [successMsg, SetSuccessMsg] = useState(""); //  Succuss Msg
  const [currentActionerEmail, SetCurrentActionerEmail] = useState("");  // current actioner email
  const [markInfoUserEmail, SetMarkInfoUserEmail] = useState([]); // mark info
  const [selectedMarkUserInfo, setSelectedMarkUserInfo] = useState(null); // selected mark info
  const [isLoading, SetIsLoading] = useState(false);  // loading state intial false
  // alert box dailog and dcesiption
  const [confirmDailogObj, setConfirmDailogObj] = useState({ Confirmtext: "", Description: "" });
  const [callbackvisible, setCallBackVisible] = useState(false);
  const [changeapprovervisible, setChangeApproverVisible] = useState(false);
  const [userComboValidationDialog, setUserComboValidationDialog] = useState(false); //  user alert validation
  const [userNotifyInfo, setUserNotifyInfo] = useState("");
  const [changeApprovercombovalue, setChangeApprovercombovalue] = useState(null); // change Approver
  const timeout = React.useRef();
  const [changeApproverhaveSecretary, setChangeApproverhaveSecretary] = useState(false); // checking change Approver have secratory
  const [statusMessage, setStatusMessage] = useState("");
  const [isPDFFullWidth, setIsPDFFullWidth] = useState(false); // change 05/04 Adding  state for PDF full width
  const [supportingDocError, setSupportingError] = useState("");
  // Passcode validation and input
  const [passcodeVerification, setPasscodeVerification] = useState(false);
  const [isPasscodeVisible, setIsPasscodeVisible] = useState(false);
  const [passcode, setPasscode] = useState("");
  const [error, setError] = useState("");
  const [validPasscode, setValidPasscode] = useState(false);
  const [validMsg, setValidmsg] = useState(false);
  // pdf component state
  const [pdfDocument, setPdfDocument] = useState(null);
  const [pdfData, setPdfData] = useState(null);

  //priti-27-04-2026
  const DEFAULT_PDF_ZOOM = 1.2;
  const FULLSCREEN_PDF_ZOOM = 2;

  const [zoomLevel, setZoomLevel] = useState(DEFAULT_PDF_ZOOM);
  const [fitToPageValue, setFitToPageValue] = useState(DEFAULT_PDF_ZOOM);
  //priti-27-04-2026

  const [gistZoomLevel, setGistZoomLevel] = useState(1); // Default zoom level
  const [currentPage, setCurrentPage] = useState(1);
  const [numPages, setNumPages] = useState(0);
  const [pages, setPages] = useState([]);
  const pdfViewerRef = useRef(null);
  const [pdfpath, setPdfPath] = useState(null);
  const [gistpdfPath, setGistPdfPath] = useState(null);
  const pageRefs = useRef([]);
  const gistPageRefs = useRef([]);
  const [verifyPasscode, setVerifypasscode] = useState("false");
  const isMobile = window.innerWidth <= 768;
  const [mobileCommentsDialog, setMobileCommentsDialog] = useState(false);
  const [mobileCommentsEditDialog, setMobileCommentsEditDialog] = useState(false);
  const [currentEditComment, setCurrentEditComment] = useState("");
  const [editIndex, setEditIndex] = useState(null);
  const [dialogEdit, setDialogEdit] = useState("");
  const [isVieweEligible, setIsVieweEligible] = useState("true");
  const [isResolutionChecked, setIsResolutionChecked] = useState(false);
  const [gistDocPopup, setGistDocPopup] = useState(false);
  const [gistPDF, setGistPdf] = useState(null);
  const [errorMessageEdit, setErrorMessageEdit] = useState("");
  const [isGistSecretary, setisGistSecretary] = useState(false);
  const [selectedATRType, setSelectedATRType] = useState(null);
  const [eNumObj, seteNumObj] = useState({});
  const [MDNameTitle, setMDNameTitle] = useState("");
  const [fileKey, setFileKey] = useState(Date.now()); // Unique key for file input
  // file upload
  // const [fileQueue, setFileQueue] = useState([]);
  const [departmentAlias, setDepartmentAlias] = useState(null); // State to store the file  
  // const [uploadedFiles, setUploadedFiles] = useState([]); // Store files for processing
  const max = 8000;
  const expansionPanelRef = useRef(null); // Ref for the expansion panel
  const [panelHeight, setPanelHeight] = useState(0);
  const [deleteNotification, setDeleteNotification] = useState(false);
  const [selectedDeleteIndex, setSelectedDeleteIndex] = useState(null);
  const [referedID, setReferedID] = useState("");
  const [showProgressBar, setShowProgressBar] = useState(false);
  const [pdfLoading, setPdfLoading] = useState(true); // true initially, will turn false after getPDFbase64 finishes
  const [progressValue, setProgressValue] = useState(0);
  const [pdfProgress, setPdfProgress] = useState(0);
  const [pdfBlob, setPdfBlob] = useState(null);
  const [actionInProgress, setActionInProgress] = useState(false);
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [gistPdfBlob, setGistPdfBlob] = useState(null);
  const [isDialogOpen, setIsDialogOpen] = useState(false);
  const [selectedComment, setSelectedComment] = useState(null);

  // Build TAT data for the grid from commentAuditLogDto
  const buildTatData = (noteData) => {
    if (!noteData || !noteData.commentAuditLogDto) return [];

    // Exclusion keywords (case-insensitive)
    const excludeActions = [
      "Gist Document Updated for Note",
      "Note Referred",
      "Note Approver Changed",
      "Note Returned",
      "Note Referred Back"
    ];

    // 1. Sort logs by sNo
    const sortedLogs = [...noteData.commentAuditLogDto].sort((a, b) => a.sNo - b.sNo);

    // 2. Filter out excluded actions FIRST
    const includedLogs = sortedLogs.filter(
      log => !excludeActions.some(keyword =>
        log.action?.toLowerCase().includes(keyword.toLowerCase())
      )
    );

    // 3. Build final TAT rows with correct Received On (previous included log)
    const tatData = includedLogs.map((log, idx) => {
      const previousLog = includedLogs[idx - 1];

      return {
        steps: simplifyActionText(log.action),
        submittedOn: previousLog ? getdateconversion(previousLog.actionDate) : getdateconversion(log.actionDate),
        approvedOn: getdateconversion(log.actionDate),
        individualTat: log?.individualTatDays ?? "-",
        officeTat: log?.officeTatDays ?? "-"
      };
    });

    return tatData;
  };

  // Helper function to simplify action text
  const simplifyActionText = (action) => {
    if (!action) return action;

    // Remove specific names and keep only the action type
    return action
      .replace(/ to [^,]+/, '') // Remove "to [Name]" 
      .replace(/ by [^,]+/, '') // Remove "by [Name]"
      .replace(/ from [^,]+/, '') // Remove "from [Name]"
      .trim();
  };
  // TAT dialog state management
  const [tatDialog, setTatDialog] = useState(false);

  const tatDialogToggle = () => {
    setTatDialog(!tatDialog);
  };

  const handleIconClick = (comment) => {
    const logs = noteData.commentAuditLogDto;

    // Find index of selected comment
    const index = logs.findIndex((l) => l.sNo === comment.sNo);

    // Previous action date (or same if first row)
    const receivedDate =
      index > 0 ? logs[index - 1].actionDate : comment.actionDate;

    // Build updated object
    const updatedComment = {
      ...comment,
      rawActionDate: comment.actionDate,
      rawReceivedDate: receivedDate
    };

    setSelectedComment(updatedComment);
    setIsDialogOpen(true);
  };

  //priti-27-04-2026
  useEffect(() => {
    const nextZoom = isPDFFullWidth ? FULLSCREEN_PDF_ZOOM : DEFAULT_PDF_ZOOM;

    setZoomLevel(nextZoom);

    requestAnimationFrame(() => {
      pdfAnnotatorRef.current?.setZoom?.(nextZoom);
    });
  }, [isPDFFullWidth]);
  //priti-27-04-2026


  useEffect(() => {
    // Update height when the expansion panel changes
    updatePanelHeight();
  }, [expanded]);

  useLayoutEffect(() => {
    // Update the panel height on page loadz
    updatePanelHeight();

    // Add a resize event listener
    const handleResize = () => {
      updatePanelHeight();
    };
    window.addEventListener("resize", handleResize);

    // Cleanup the event listener on component unmount
    return () => {
      window.removeEventListener("resize", handleResize);
    };
  }, []);

  const updatePanelHeight = () => {
    requestAnimationFrame(() => {
      if (expansionPanelRef.current && pdfViewerRef.current) {
        const viewer = pdfViewerRef.current;

        if (isPDFFullWidth) {
          // FULL SCREEN MODE
          viewer.style.height = "100vh";
          viewer.style.overflowY = "visible";
          viewer.style.overflowX = "visible";
          viewer.style.border = "none";
        } else {
          if (window.innerWidth > 768) {
            // DESKTOP NON-FULLSCREEN MODE
            const expansionPanelHeight = expansionPanelRef.current.offsetHeight || 0;
            const adjustedHeight = expansionPanelHeight > 0 ? expansionPanelHeight : 0;
            setPanelHeight(adjustedHeight);
            viewer.style.height = `${adjustedHeight - 30}px`; // -30px for padding/margin
          } else {
            // MOBILE NON-FULLSCREEN MODE
            const fixedMobileHeight = 500;
            viewer.style.height = `${fixedMobileHeight}px`;
          }

          // Restore normal styling
          viewer.style.overflowY = "auto";
          viewer.style.overflowX = "auto";
          viewer.style.border = "1px solid #00000030";
        }
      }
    });
  };

  const handlePanelToggle = (panelId) => {
    setExpanded((prevExpanded) => {
      const newExpanded = prevExpanded.includes(panelId)
        ? prevExpanded.filter((id) => id !== panelId)
        : [...prevExpanded, panelId];
      return newExpanded;
    });
  };

  // Add a MutationObserver to detect changes in the expansion panel's height
  useEffect(() => {
    const observer = new MutationObserver(() => {
      updatePanelHeight();
    });

    if (expansionPanelRef.current) {
      observer.observe(expansionPanelRef.current, {
        attributes: true,
        childList: true,
        subtree: true,
      });
    }

    // Cleanup observer on component unmount
    return () => {
      if (observer) observer.disconnect();
    };
  }, []);


  useEffect(() => {
    const MDname = "MD & CEO";
    setMDNameTitle(MDname);
    isUserViewEligible();
    VerifyUserPasscode();

    const updateProgress = (value) => {
      setProgressValue((prev) => (value > prev ? value : prev));
    };

    const fetchData = async () => {
      try {
        setShowProgressBar(true);
        setProgressValue(0);
        updateProgress(5);

        const accessToken = await getAccessToken({ ...loginRequest, account }, instance);
        updateProgress(10);

        // Fire getPDFbase64 in parallel (not tracked in progress bar)
        getPDFbase64(accessToken); // Do not await

        // === Start Parallel API Calls ===
        const generalDetailsPromise = fetch(
          `${API_BASE_URL}${API_ENDPOINTS.eNote_GetGeneralDetails}`,
          {
            method: "POST",
            headers: {
              ...API_COMMON_HEADERS,
              Authorization: `Bearer ${accessToken}`,
            },
            body: JSON.stringify({ noteId: id.id }),
          }
        );

        const atrCreatorsPromise = fetch(
          `${API_BASE_URL}${API_ENDPOINTS.eNote_GetATRCreators}`,
          {
            method: "GET",
            headers: {
              ...API_COMMON_HEADERS,
              Authorization: `Bearer ${accessToken}`,
            },
          }
        );

        const dropdownDataPromise = fetch(
          `${API_BASE_URL}${API_ENDPOINTS.GET_DROPDOWNDATA}`,
          {
            method: "GET",
            headers: {
              ...API_COMMON_HEADERS,
              Authorization: `Bearer ${accessToken}`,
            },
          }
        );

        const [generalRes, atrCreatorsRes, dropdownRes] = await Promise.all([
          generalDetailsPromise,
          atrCreatorsPromise,
          dropdownDataPromise,
        ]);

        const [data, aTRCreatoruserInfo, enumsObj] = await Promise.all([
          generalRes.json(),
          atrCreatorsRes.json(),
          dropdownRes.json(),
        ]);

        updateProgress(30);

        setNoteData(data);
        setDepartmentAlias(data.departmentAlias);
        SetCurrentActionerEmail(data.currentActioner);
        seteNumObj(enumsObj);
        SetATRCreaters(aTRCreatoruserInfo);

        const currentUserObj = data.noteApproversDTO?.filter(
          (obj) => obj.approverEmail === data.currentActioner
        );

        const getNoteSecretaryPromise = fetch(
          `${API_BASE_URL}${API_ENDPOINTS.eNote_GetNoteSecretary}`,
          {
            method: "POST",
            headers: {
              ...API_COMMON_HEADERS,
              Authorization: `Bearer ${accessToken}`,
            },
            body: JSON.stringify({
              noteId: id.id,
              noteApproverId: currentUserObj[0]?.noteApproverId,
            }),
          }
        );

        const getNoteSecretaryRes = await getNoteSecretaryPromise;
        const getNoteSecretaryUserArray = await getNoteSecretaryRes.json();
        updateProgress(45);

        const markforinfo = (data?.noteMarkedInfoDTO || []).map((obj) => ({
          noteId: obj.noteId,
          markedEmail: obj.markedEmail,
          createdBy: obj.createdBy,
          markedEmailName: obj.markedEmailName,
        }));
        SetMarkInfoUserEmail(markforinfo);

        const referredUserInfo = data.noteReferrerDTO || [];
        const isReferred = referredUserInfo.some(
          (obj) =>
            obj.referrerEmail === data.currentActioner &&
            obj.referrerStatus === 1 &&
            obj.referrerEmail === accounts[0].username
        );
        SetIsreferredUser(isReferred);

        const getNoteSecretaryUserTempArray = [...getNoteSecretaryUserArray];
        const crntApvrObj = data.noteApproversDTO?.filter(
          (obj) => obj.approverEmail === accounts[0].username && obj.approverType === 2
        );
        const isCrntSecretary = data?.noteSecretaryDTO?.some(
          (obj) =>
            obj.secretaryEmail === accounts[0].username ||
            obj.approverEmail === crntApvrObj[0]?.approverEmail
        );

        setIsSecretary(isCrntSecretary);
        setGetNoteSecretaryUserInfo(getNoteSecretaryUserTempArray);
        updateProgress(65);

        const currentApproverObj = data?.noteApproversDTO?.filter(
          (_x) => _x.approverEmail === accounts[0].username && _x.approverType === 2
        );

        const isATRCreator =
          currentApproverObj.length > 0 &&
          aTRCreatoruserInfo?.some(
            (obj) => obj.atrCreatorEmail === currentApproverObj[0]?.approverEmail
          );
        setIsCurrentATRCreator(isATRCreator);

        if (isATRCreator) {
          const refAccessToken = await getAccessToken({ ...loginRequest, account }, instance);
          const params = { ATRCreatorEmail: accounts[0].username };

          try {
            const assigneeRes = await fetch(
              `${API_BASE_URL}${API_ENDPOINTS.eNote_GetATRAssignees}`,
              {
                method: "POST",
                body: JSON.stringify(params),
                headers: {
                  ...API_COMMON_HEADERS,
                  Authorization: `Bearer ${refAccessToken}`,
                },
              }
            );
            const assignees = await assigneeRes.json();
            setAssigneeUsers(assignees);
            setAssigneerUserDetails(assignees);
          } catch (err) {
            console.error("Assignees API error", err);
          }
        }

        updateProgress(75);

        await Promise.all(
          (data?.noteSecretaryDTO || []).map(async (obj) => {
            if (
              obj.secretaryEmail === accounts[0].username ||
              obj.approverEmail === accounts[0].username
            ) {
              const isValidPath = obj.gistWordDocumentPath && obj.gistWordDocumentPath.length > 0;
              const fileExtension = obj.gistWordDocumentFileName?.split(".").pop().toLowerCase();
              const hasGistPdfDocumentPath = isValidPath && fileExtension === "pdf";

              const buttonText = isApprover() ? "Noted" : "Approve";
              const isApproveButtonVisible = buttonText === "Approve";

              if (hasGistPdfDocumentPath && isApproveButtonVisible) {
                setGistDocPopup(true);
                setGistPdf(obj.gistWordDocumentPath);

                const token = await getAccessToken({ ...loginRequest, account }, instance);
                const params = { supportingDocumentPath: obj.gistWordDocumentPath };

                const getSupportingDocs = await fetch(
                  `${API_BASE_URL}${API_ENDPOINTS.GET_Base64}`,
                  {
                    method: "POST",
                    headers: {
                      ...API_COMMON_HEADERS,
                      Authorization: `Bearer ${token}`,
                    },
                    body: JSON.stringify(params),
                  }
                );

                const supportDocDetails = await getSupportingDocs.text();
                if (supportDocDetails !== "File Not Available") {
                  renderGistPDF(supportDocDetails);
                }
              } else {
                setGistDocPopup(false);
              }
            }
          })
        );

        updateProgress(90);
        ValidationForExpendpanel(isATRCreator, isCrntSecretary, data, enumsObj);
        getAttachGistDoc(data);
        updateProgress(100);

        setTimeout(() => setShowProgressBar(false), 400);
        // SetIsLoading(false);

      } catch (error) {
        console.error("Error in fetchData", error);
        setShowProgressBar(false);
      }
    };

    fetchData();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Will convert base64 to pdf -- Kavya (11/07)
  const renderGistPDF = (supportDocDetails) => {
    setGistPdfPath(supportDocDetails);
    try {
      const binaryString = window.atob(supportDocDetails);
      const len = binaryString.length;
      const bytes = new Uint8Array(len);
      for (let i = 0; i < len; ++i) {
        bytes[i] = binaryString.charCodeAt(i);
      }

      const blob = new Blob([bytes], { type: 'application/pdf' });
      setGistPdfBlob(blob);

    } catch (error) {
      console.error("Error creating PDF Blob:", error);
    }
  };

  // validation for page view
  const isUserViewEligible = async () => {
    let isUserViewEligible = true;

    try {
      const accessToken = await getAccessToken(
        { ...loginRequest, account },
        instance
      );
      const response = await fetch(
        `${API_BASE_URL}${API_ENDPOINTS.GET_NoteViewEligibility}?userPrincipalName=${accounts[0].username}&NoteId=${id.id}`,
        {
          method: "POST",
          headers: {
            ...API_COMMON_HEADERS,
            Authorization: `Bearer ${accessToken}`,
          },
        }
      );

      const result = await response.text();
      if (result) {
        isUserViewEligible = result;
        if (result === "false") {
          // SetIsLoading(false);
        }

        setIsVieweEligible(result);
      }
      return isUserViewEligible;
    } catch (error) {
      // SetIsLoading(false);
      console.error("Error fetching data:", error);
    }
  };



  // In your async function replace setPdfProgress calls with setProgressAnimated

  let progressTimer = null;

  const startProgressBar = () => {
    // Start incrementing progressValue from 1 to 99
    setPdfProgress(1);
    if (progressTimer) clearInterval(progressTimer);
    progressTimer = setInterval(() => {
      setPdfProgress((prev) => {
        if (prev < 99) return prev + 1;
        return prev;
      });
    }, 200); // Adjust speed as needed
  };

  const stopProgressBar = () => {
    if (progressTimer) clearInterval(progressTimer);
    setPdfProgress(100);
  };

  const getPDFbase64 = async (accessToken) => {
    setPdfLoading(true);
    startProgressBar();
    try {
      const fileUrl = `${API_BASE_URL}${API_ENDPOINTS.GET_Base64PDF}?noteId=${id.id}`;
      const chunkSizeMB = 5; // 5MB per chunk
      const concurrency = 4; // 4 parallel downloads

      // 1. Get file size
      const headRes = await fetch(fileUrl, {
        method: "GET",
        headers: {
          ...API_COMMON_HEADERS,
          Authorization: `Bearer ${accessToken}`,
        },
      });
      const fileSize = parseInt(headRes.headers.get("content-length"), 10);
      if (!fileSize) throw new Error("Could not determine file size.");

      const chunkSize = chunkSizeMB * 1024 * 1024;
      const numChunks = Math.ceil(fileSize / chunkSize);

      // 2. Create chunk download tasks
      const chunks = Array.from({ length: numChunks }, (_, i) => ({
        index: i,
        start: i * chunkSize,
        end: Math.min((i + 1) * chunkSize - 1, fileSize - 1),
        retries: 0,
      }));

      // 3. Download chunks with limited concurrency
      const results = new Array(numChunks);
      let completed = 0;

      const downloadChunk = async (chunk) => {
        try {
          const res = await fetch(fileUrl, {
            headers: {
              ...API_COMMON_HEADERS,
              Authorization: `Bearer ${accessToken}`,
              Range: `bytes=${chunk.start}-${chunk.end}`,
            },
          });
          if (!res.ok && res.status !== 206) throw new Error(`Chunk ${chunk.index} failed`);
          const buffer = await res.arrayBuffer();
          results[chunk.index] = buffer;
          completed++;
          setPdfProgress(Math.round((completed / numChunks) * 100));
        } catch (err) {
          if (chunk.retries < 3) {
            chunk.retries++;
            await downloadChunk(chunk);
          } else {
            throw err;
          }
        }
      };

      // 4. Parallel chunk download controller
      const queue = [...chunks];
      const workers = Array(concurrency).fill(null).map(async () => {
        while (queue.length) {
          const chunk = queue.shift();
          if (chunk) await downloadChunk(chunk);
        }
      });
      await Promise.all(workers);

      // 5. Combine chunks into single Blob
      const pdfBlob = new Blob(results, { type: "application/pdf" });
      setPdfBlob(pdfBlob);
      console.log(pdfBlob, "pdfBlob");

      // 6. Pass the Blob to PDF.js directly
      const fileReader = new FileReader();
      fileReader.onloadend = async () => {
        const bytes = new Uint8Array(fileReader.result);
        setPdfData(bytes);
        setPdfLoading(false);
        stopProgressBar();
      };
      fileReader.readAsArrayBuffer(pdfBlob);
    } catch (error) {
      console.error("Error fetching or rendering PDF:", error);
      setPdfProgress(0);
      setPdfLoading(false);
      stopProgressBar();
    }
  };

  // Zoom in the file -- Kavya (12/07)
  const downloadBase64PDFFile = async (path) => {
    try {
      SetIsLoading(true);
      // Get access token
      const accessToken = await getAccessToken(
        { ...loginRequest, account },
        instance
      );

      // API request
      const response = await fetch(`${API_BASE_URL}${API_ENDPOINTS.GET_Base64}`, {
        method: "POST",
        headers: {
          ...API_COMMON_HEADERS,
          Authorization: `Bearer ${accessToken}`,
        },
        body: JSON.stringify({ supportingDocumentPath: path }),
      });

      const pdfDetails = await response.text();

      // Handle "File Not Available!" response
      if (!response.ok || pdfDetails === "File Not Available!") {
        throw new Error("File Not Available!");
      }

      // Convert Base64 to ByteArray
      const byteArray = new Uint8Array(
        atob(pdfDetails).split("").map((c) => c.charCodeAt(0))
      );

      // Create a Blob and Object URL
      const blob = new Blob([byteArray], { type: "application/pdf" });
      const url = URL.createObjectURL(blob);

      // Create and trigger a download link
      const link = document.createElement("a");
      link.href = url;
      link.download = noteData?.notePdfFileName || "download.pdf"; // Fallback filename
      document.body.appendChild(link);
      link.click();

      // Clean up
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
    } catch (error) {
      console.error("Error downloading PDF:", error.message);
    } finally {
      SetIsLoading(false);
    }
  };

  // download word file
  const downloadBase64WordFile = async (path, fileName) => {
    SetIsLoading(true);
    const accessToken = await getAccessToken(
      { ...loginRequest, account },
      instance
    );
    const parmas = {
      supportingDocumentPath: path,
    };
    const wordDoc = await fetch(`${API_BASE_URL}${API_ENDPOINTS.GET_Base64}`, {
      method: "POST",
      headers: {
        ...API_COMMON_HEADERS,
        Authorization: `Bearer ${accessToken}`,
      },
      body: JSON.stringify(parmas),
    });
    const wordDocDetails = await wordDoc?.text();
    if (wordDocDetails !== "File Not Available!") {
      // Convert base64 to binary
      const byteCharacters = atob(wordDocDetails);
      const byteNumbers = new Array(byteCharacters.length);
      for (let i = 0; i < byteCharacters.length; i++) {
        byteNumbers[i] = byteCharacters.charCodeAt(i);
      }
      const byteArray = new Uint8Array(byteNumbers);
      // Create a Blob object
      const blob = new Blob([byteArray], { type: "application/doc" });
      const link = document.createElement("a");
      link.href = URL.createObjectURL(blob);
      link.download = fileName;
      document.body.appendChild(link);
      link.click();

      // Clean up and remove the link
      document.body.removeChild(link);
      URL.revokeObjectURL(link.href);
    }
    SetIsLoading(false);
  };

  // Get Gist Doc info -coverting base64 to url 
  const getAttachGistDoc = (data) => {
    // checking login user = secretaryEmail or approverEmail ,displaying  doc for login user
    data?.noteSecretaryDTO?.filter(obj => {
      if (obj.secretaryEmail === accounts[0].username ||
        obj.approverEmail === accounts[0].username) {
        if (obj.gistWordDocumentFileName !== null &&
          obj.gistWordDocumentPath !== null) {
          getConvertingGistDocBase64ToUrl(obj.gistWordDocumentPath, obj.gistWordDocumentFileName);
        }
      }
    });
  }
  // download supporting doc
  const getSupportDocHyperlink = async (filePath, fileName) => {
    SetIsLoading(true);
    const accessToken = await getAccessToken(
      { ...loginRequest, account },
      instance
    );

    const parmas = {
      supportingDocumentPath: filePath,
    };
    const getSupportingDocs = await fetch(
      `${API_BASE_URL}${API_ENDPOINTS.GET_Base64}`,
      {
        method: "POST",
        headers: {
          ...API_COMMON_HEADERS,
          Authorization: `Bearer ${accessToken}`,
        },
        body: JSON.stringify(parmas),
      }
    );
    const supportDocDetails = await getSupportingDocs?.text();

    if (supportDocDetails !== "File Not Available!") {
      const byteCharacters = atob(supportDocDetails);
      const byteNumbers = new Array(byteCharacters.length);
      for (let i = 0; i < byteCharacters.length; i++) {
        byteNumbers[i] = byteCharacters.charCodeAt(i);
      }
      // Create a Blob object
      const byteArray = new Uint8Array(byteNumbers);
      // checking file type
      const fileType = getFileType(fileName.toLowerCase());
      // Create a Blob object
      const blob = new Blob([byteArray], { type: `application/${fileType}` });

      const link = document.createElement("a");
      link.href = URL.createObjectURL(blob);
      link.download = fileName;
      document.body.appendChild(link);
      link.click();

      // Clean up and remove the link
      document.body.removeChild(link);
      URL.revokeObjectURL(link.href);
    }
    SetIsLoading(false);
  };

  //converting base64 to url - gistDoc
  const getConvertingGistDocBase64ToUrl = async (filePath, fileName) => {
    let base64Url = "";
    const accessToken = await getAccessToken(
      { ...loginRequest, account },
      instance
    );

    const parmas = {
      supportingDocumentPath: filePath,
    };
    const getSupportingDocs = await fetch(
      `${API_BASE_URL}${API_ENDPOINTS.GET_Base64}`,
      {
        method: "POST",
        headers: {
          ...API_COMMON_HEADERS,
          Authorization: `Bearer ${accessToken}`,
        },
        body: JSON.stringify(parmas),
      }
    );
    const supportDocDetails = await getSupportingDocs?.text();

    if (supportDocDetails !== "File Not Available!") {
      const byteCharacters = atob(supportDocDetails);
      const byteNumbers = new Array(byteCharacters.length);
      for (let i = 0; i < byteCharacters.length; i++) {
        byteNumbers[i] = byteCharacters.charCodeAt(i);
      }
      // Create a Blob object
      const byteArray = new Uint8Array(byteNumbers);
      // checking file type
      const fileType = getFileType(fileName.toLowerCase());
      const blob = new Blob([byteArray], { type: `application/${fileType}` });
      base64Url = URL.createObjectURL(blob);
      SetWordPDFInfowarring({
        ...wordandPdfInfo,
        wordInfo: {
          fileExtension: "",
          fileName: fileName === null ? "" : fileName,
          filePath: "",
          isValid: false,
          isDownloadble: true,
          base64: base64Url,
        },
      });
    }
    return base64Url;
  };

  // getfileType  -checking file type
  const getFileType = (filename) => {
    let fileType = "doc";
    if (filename.endsWith(".pdf")) {
      fileType = "pdf";
    }
    if (filename.endsWith(".doc") || filename.endsWith(".docx")) {
      fileType = "doc";
    }
    if (
      filename.endsWith(".png") ||
      filename.endsWith(".jpg") ||
      filename.endsWith(".img") ||
      filename.endsWith(".svg")
    ) {
      fileType = "image";
    }
    if (filename.endsWith(".txt")) {
      fileType = "txt";
    }
    if (filename.endsWith(".xlsx")) {
      fileType = "xlsx";
    }
    if (filename.endsWith(".eml")) {
      fileType = "eml";
    }
    return fileType;
  };
  // expansion panel
  const ValidationForExpendpanel = (isATRCreator, isCurrentSecretary, data, enumsObj) => {
    console.log(data, "data")
    // Precompute gistSecretaries and isGistSecretary outside the mapping function
    const gistSecretaries = data?.gistSecretaries
      ? data.gistSecretaries.split(',').map(email => email.trim())
      : [];

    const isGistSecretary = gistSecretaries?.some(email => email === accounts[0].username);
    setisGistSecretary(isGistSecretary);

    const updatedExpend = expendJson?.map((obj) => {
      if (isATRCreator && obj.id === "ATR Assignees") {
        return {
          ...obj,
          isVisible: !obj.isVisible,
        };
      } else if (
        (isCurrentSecretary || isGistSecretary) &&  // Show if either is true
        obj.id === "Gist Document"
      ) {
        // Update visibility for both isCurrentSecretary and gistSecretaries
        return {
          ...obj,
          isVisible: !obj.isVisible,
        };
      } else if (
        (obj.id === "General Comments" || obj.id === "Attach Supporting Documents") &&
        (data?.status === 2 || data?.status === 3 || data?.status === 9) &&
        (data?.noteApproversDTO?.some(
          approver =>
            approver.approverEmail === accounts[0].username &&
            approver.approverEmail === data?.currentActioner
        ) ||
          data?.noteReferrerDTO?.some(
            referrer =>
              referrer.referrerEmail === accounts[0].username &&
              referrer.referrerEmail === data?.currentActioner
          ))
      ) {
        return {
          ...obj,
          isVisible: !obj.isVisible,
        };
      } else if (
        obj.id === "Mark for Information Section" &&
        data?.status === 5 &&
        data?.createdBy === accounts[0]?.username
      ) {
        return {
          ...obj,
          isVisible: !obj.isVisible,
        };
      } else {
        return obj; // Return the original object if no conditions matched
      }
    });

    SetExpendJson(updatedExpend);
    return updatedExpend;
  };

  // all uploading attchments controller function start here
  const noteWordconvertBase64Word = () => {
    var selectedFile = document.getElementById("noteWordDocfile").files;
    let cstWarningMsg = "";
    let isValidFile = true;
    if (selectedFile.length > 0) {
      var fileName = selectedFile[0].name;

      if (
        !(
          fileName.toLowerCase().endsWith(".docx") ||
          fileName.toLowerCase().endsWith("doc")
        )
      ) {
        cstWarningMsg = "File type not allowed";
        isValidFile = false;
      }

      if (selectedFile[0].size > 10485760) {
        cstWarningMsg = "File size should not exceed more then 10 MB";
        isValidFile = false;
      }
      if (checkingSpecialCharPattern(selectedFile[0].name)) {
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
          const RemovedHeaderFormBase64 = base64.split(",");
          setNoteWordDoc({
            noteWordDocInfo: {
              fileName: selectedFile[0].name,
              filePath: RemovedHeaderFormBase64[1],
              isValid: isValidFile,
              warningMsg: cstWarningMsg,
            },
          });
        };
        // Convert data to base64
        fileReader.readAsDataURL(fileToLoad);
      } else {
        setNoteWordDoc({
          noteWordDocInfo: {
            fileName: selectedFile[0].name,
            filePath: "",
            isValid: isValidFile,
            warningMsg: cstWarningMsg,
          },
        });
      }
    }
  };

  const multipleDocUpload = async () => {
    // Added passcode redirect page --> Kavya (23/07)
    if (verifyPasscode === "true") {
      let reamingFileSize = 26214400;
      const inValidFileNames = Object.keys(supportDocWarning).filter(
        (fileName) => supportDocWarning[fileName]?.isValid === false
      );
      // filter valid files only if()
      if (supportDocfilesInfo.length > 0) {
        const vaildMultiplefile = supportDocfilesInfo.filter((obj) => {
          if (!inValidFileNames.includes(obj.supportingDocumentFileName)) {
            return obj;
          }
        });
        vaildMultiplefile.map((obj) => {
          reamingFileSize = reamingFileSize - obj.supportingDocumentPathLength;
        });
      }
      if (noteData?.noteSupportingDocumentsDTO.length > 0) {
        noteData?.noteSupportingDocumentsDTO.map((obj) => {
          reamingFileSize = reamingFileSize - obj.supportingDocumentPathLength;
        });
      }
      // }
      let warningmsg = "";
      // eslint-disable-next-line no-useless-escape
      const validname = /[!@#$%^&*(){}\[\];:,<>\?\/\\]/;
      const selectedFile = document.getElementById("multiDoc").files;
      const TempfileInfo = supportDocfilesInfo;
      const promises = [];
      const fileCount = selectedFile.length;
      for (let i = 0; i < fileCount; i++) {
        const fileToLoad = selectedFile[i];
        const fileExtession = fileToLoad.name.split(".");
        // Check for special characters in the filename

        if (
          !noteData?.noteSupportingDocumentsDTO?.some(
            (obj) => obj.supportingDocumentFileName === fileToLoad.name
          )
        ) {
          if (
            !(
              (
                fileToLoad.name.toLowerCase().endsWith(".docx") ||
                fileToLoad.name.toLowerCase().endsWith(".doc") ||
                fileToLoad.name.toLowerCase().endsWith(".pdf") ||
                fileToLoad.name.toLowerCase().endsWith(".xlsx")
              )
            )
          ) {
            warningmsg = "File type is not allowed";
            updateSupportDocWarning(
              fileToLoad.name,
              warningmsg,
              fileExtession[fileExtession.length - 1]
            );
          }
          if (validname.test(fileToLoad.name)) {
            warningmsg = "File name should not contain special characters";
            updateSupportDocWarning(
              fileToLoad.name,
              warningmsg,
              fileExtession[fileExtession.length - 1]
            );
          }
          if (fileToLoad.size > reamingFileSize) {
            warningmsg =
              "Cumulative size of all the supporting documents should not be exceeded 25 MB.";
            setSupportingError(warningmsg);
          }

          // File Reader.....
          const fileReader = new FileReader();
          const promise = new Promise((resolve) => {
            fileReader.onload = function (fileLoadedEvent) {
              // Added for base 64 split params ---> Kavya(18-07)
              const base64 = fileLoadedEvent.target.result.split(",")[1];
              const partLength = Math.ceil(base64.length / 10);
              const parts = [];
              for (let j = 0; j < 10; j++) {
                parts.push(base64.slice(j * partLength, (j + 1) * partLength));
              }

              resolve({
                name: fileToLoad.name,
                parts: parts,
                size: fileToLoad.size,
              });
            };
          });

          fileReader.readAsDataURL(fileToLoad);
          promises.push(promise);
        }
      }

      Promise.all(promises)
        .then((fileDataArray) => {
          const updatedTempFileInfo = fileDataArray.reduce(
            (acc, fileData) => {
              // const decodedfile =atob

              const ObjExist = acc.map((obj) => obj.supportingDocumentFileName);
              if (!ObjExist.includes(fileData.name)) {
                acc.push({
                  // noteSupportingDocumentId: 0,
                  noteId: noteData.noteId,
                  // supportingDocumentPath: fileData.base64.split(",")[1],
                  supportingDocumentFileName: fileData.name,
                  createdDate: new Date(),
                  createdBy: accounts[0].username,
                  modifiedDate: new Date(),
                  modifiedBy: accounts[0].username,
                  // Added for base 64 split params parts---> Kavya(18-07)
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
                  supportingDocumentPathLength: fileData.size,
                });
              }
              return acc;
            },
            [...TempfileInfo]
          );
          setSupportDocfilesInfo(updatedTempFileInfo);
        })

        .catch((error) => {
          console.error("Error reading files:", error);
        });
    } else {
      setValidPasscode(true);
      // Changed the content --> Kavya (25-07)
      setValidmsg(
        "Passcode is not set. Please create passcode to proceed further."
      );
    }
  };

  const attchemtsForGistDoc = () => {
    var selectedFile = document.getElementById("WordDocfile").files;
    let waringmsg = "";
    let isValidFile = true;
    if (selectedFile.length > 0) {
      var fileExtension = selectedFile[0].name.split(".");
      const nameFile = selectedFile[0].name;
      if (
        !(
          nameFile.toLowerCase().endsWith(".pdf")
        )
      ) {
        waringmsg = "File type not allowed";
        SetfilesWordWarning(waringmsg);
        isValidFile = false;
        SetWordPDFInfowarring({
          ...wordandPdfInfo,
          wordInfo: {
            fileExtension: `.${fileExtension[fileExtension.length - 1]}`,
            fileName: selectedFile[0].name,
            filePath: "",
            isValid: false,
            isDownloadble: false,
            base64: ""
          },
        });

        return isValidFile;
      }
      /* 5MB = 5242880  */
      if (selectedFile[0].size > 5242880) {
        waringmsg = "File size should not exceed more then 5 MB";
        isValidFile = false;
        SetfilesWordWarning(waringmsg);
        SetWordPDFInfowarring({
          ...wordandPdfInfo,
          wordInfo: {
            fileExtension: `.${fileExtension[fileExtension.length - 1]}`,
            fileName: selectedFile[0].name,
            filePath: "",
            isValid: false,
            isDownloadble: false,
            base64: ""
          },
        });
        return isValidFile;
      }
      if (checkingSpl(selectedFile[0].name)) {
        isValidFile = false;
        waringmsg = "File name sholud not contain special characters";
        SetfilesWordWarning(waringmsg);
        SetWordPDFInfowarring({
          ...wordandPdfInfo,
          wordInfo: {
            fileExtension: `.${fileExtension[fileExtension.length - 1]}`,
            fileName: selectedFile[0].name,
            filePath: "",
            isValid: false,
            isDownloadble: false,
            base64: ""
          },
        });
        return isValidFile;
      }
      if (isValidFile === true) {
        var fileToLoad = selectedFile[0];
        // FileReader function for read the file.
        var fileReader = new FileReader();
        // var base64;
        // Onload of file read the file content
        fileReader.onload = function (fileLoadedEvent) {
          let base64 = fileLoadedEvent.target.result;
          const RemovedHeaderFormBase64 = base64.split(",")[1];

          const partLength = Math.ceil(RemovedHeaderFormBase64.length / 10);
          const parts = [];

          for (let i = 0; i < 10; i++) {
            parts.push(RemovedHeaderFormBase64.slice(i * partLength, (i + 1) * partLength));
          }
          SetfilesWordWarning("");
          SetWordPDFInfowarring({
            ...wordandPdfInfo,
            wordInfo: {
              fileExtension: `.${fileExtension[fileExtension.length - 1]}`,
              fileName: selectedFile[0].name,
              // filePath: RemovedHeaderFormBase64[1],
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
              isValid: true,
              isDownloadble: false,
              base64: ""
            },
          });
        };
        // Convert data to base64
        fileReader.readAsDataURL(fileToLoad);
      }
    }
  };

  const updateSupportDocWarning = (fileName, warningMsg, fileExtension) => {
    SetSupportDocWarning((prevState) => ({
      ...prevState,
      [fileName]: {
        filename: fileName,
        isValid: false,
        warningMsg: warningMsg,
        fileExtnsion: fileExtension,
      },
    }));
    setFileKey(Date.now());
  };

  const checkingSpl = (test) => {
    var specialCharPattern = /[!@#$%^&*(){}\[\];:,<>\?\/\\]/;

    // Test the input string against the pattern
    return specialCharPattern.test(test);
  };

  // checking file type and return file icon
  const getFileIcon = (fname) => {
    var filename = fname.toLowerCase();
    let fileType = fileIcon;
    if (filename.endsWith(".pdf")) {
      fileType = filePdfIcon;
    }
    if (filename.endsWith(".doc") || filename.endsWith(".docx")) {
      fileType = fileWordIcon;
    }
    if (
      filename.endsWith(".png") ||
      filename.endsWith(".jpg") ||
      filename.endsWith(".img") ||
      filename.endsWith(".svg")
    ) {
      fileType = fileImageIcon;
    }
    if (filename.endsWith(".txt")) {
      fileType = fileTxtIcon;
    }
    if (filename.endsWith(".xlsx")) {
      fileType = fileDataIcon;
    }
    return fileType;
  };

  // checking on specialCharPattern char inn file name
  const checkingSpecialCharPattern = (test) => {
    // eslint-disable-next-line no-useless-escape
    var specialCharPattern = /[!@#$%^&*(){}\[\];:,<>\?\/\\]/;
    // Test the input string against the pattern
    return specialCharPattern.test(test);
  };

  // remove multiple attchemnts suppoutDoc
  const onRemoveMultiAttachment = (id) => {
    const filename = supportDocfilesInfo.find(
      (obj, index) => index === id
    ).supportingDocumentFileName;
    delete supportDocWarning[filename];
    SetSupportDocWarning(supportDocWarning);
    // let reamingFileSize = 26214400;
    let totalFileSize = 0;
    const inValidFileNames = Object.keys(supportDocWarning).filter(
      (fileName) => supportDocWarning[fileName]?.isValid === false
    );
    // filter valid files only if()
    if (supportDocfilesInfo.length > 0) {
      const vaildMultiplefile = supportDocfilesInfo.filter((obj) => {
        if (!inValidFileNames.includes(obj.supportingDocumentFileName)) {
          return obj;
        }
      });
      vaildMultiplefile.map((obj, index) => {
        if (index !== id) {
          totalFileSize = totalFileSize + obj.supportingDocumentPathLength;
        }
      });
    }

    // 11/06
    if (noteData?.noteSupportingDocumentsDTO?.length > 0) {
      noteData?.noteSupportingDocumentsDTO.map((obj) => {
        totalFileSize = totalFileSize + obj.supportingDocumentPathLength;
      });
    }
    if (totalFileSize <= 26214400) {
      setSupportingError("");
    }

    setSupportDocfilesInfo(
      supportDocfilesInfo.filter((obj, ind) => ind !== id)
    );
    setFileKey(Date.now());
  };

  // remove attacment for gist
  const onRemoveAttachmentWarning = (key) => {
    SetWordPDFInfowarring({
      ...wordandPdfInfo,
      [key]: {
        fileExtension: "",
        fileName: "",
        filePath: "",
        isValid: false,
        isDownloadble: false,
        base64: "",
      },
    });
    setFileKey(Date.now());
  };

  // removeNoteDoc
  const onNoteWordDocRemoveAttachmentWarning = () => {
    setNoteWordDoc({
      noteWordDocInfo: {
        fileName: "",
        warningMsg: "",
        filePath: "",
        isValid: false,
      },
    });
  };
  // all attchments controller function closed here

  // all comobox functionalities start here, on selecting  change approver -this combobox is availble in Requester
  const handleChangeApproverCombo = (e) => {
    const { value } = e.target;
    let isSecretaryExist = false;
    if (value) {
      const isApprover = noteData?.noteApproversDTO?.some(
        (obj) => obj.approverEmail === value.userPrincipalName
      );
      const isreferredUser = noteData?.noteReferrerDTO?.some(
        (obj) => obj.referredEmail === value.userPrincipalName
      );
      if (
        value.userPrincipalName === noteData?.createdBy ||
        value.userPrincipalName === noteData.currentActioner ||
        value.userPrincipalName === accounts[0].username ||
        isApprover ||
        isreferredUser
      ) {
        setDialogForRefereeAndChangeApprover(false);
        setUserNotifyInfo(
          "The selected approver cannont be same as existing Reviewers/Requester/referee/CurrentActioner"
        );
        setUserComboValidationDialog(true);
        setChangeApprovercombovalue("");
      } else {
        noteData?.noteApproversDTO?.filter(async (obj) => {
          if (
            obj.approverEmail === noteData?.currentActioner &&
            obj.approverType === 2
          ) {
            isSecretaryExist = await findSecrateries(obj.approverType, value);
            if (
              isSecretaryExist &&
              noteData?.noteWordFileName === null &&
              noteData?.noteWordPath === null
            ) {
              setChangeApproverhaveSecretary(true);
            }
          }
        });
        setChangeApprovercombovalue(value);
      }
    } else {
      setChangeApprovercombovalue(value);
      setChangeApproverhaveSecretary(false);
      setNoteWordDoc({
        noteWordDocInfo: {
          fileName: "",
          filePath: "",
          isValid: false,
          warningMsg: "",
        },
      });
    }
  };

  // Refer combobox -Selected Refer user value store in state
  const selectedReferrUser = (event) => {
    SetSelectedRefer(event.value);
  };
  // MarkInfo user combobox -selected user value stote in state
  const onchangehanleMarkInfo = (event) => {
    setSelectedMarkUserInfo(event.value);
  };

  // add assignee /mark info user details  in table selected markinfo user details adding in table

  // remove assignees /mark info user details in table
  const onDeleteAssigneeUser = (ind) => {
    setSelectedAssigneeUsersinfo(
      selectedAssigneeUsersinfo.filter((obj, index) => index !== ind)
    );
  };

  //Find secretary exists for the added Approvers/Reviewers - this function will call  while Requester is chnageing approver
  const findSecrateries = async (approverType, user) => {
    let param = {
      noteApproversMasterlst: [
        { approverType: approverType, approverEmail: user.userPrincipalName },
      ],
      departmentName: user.department,
    };
    const accessToken = await getAccessToken(
      { ...loginRequest, account },
      instance
    );

    let secretaryExists = false;

    try {
      await fetch(
        `${API_BASE_URL}${API_ENDPOINTS.eNote_GetNoteSecretaryofApprover}`,
        {
          method: "POST",
          headers: {
            ...API_COMMON_HEADERS,
            Authorization: `Bearer ${accessToken}`,
          },
          body: JSON.stringify(param),
        }
      )
        .then((response) => {
          return response.json();
        })
        .then((data) => {
          secretaryExists = data.secretaryExist;
        });
    } catch (error) {
      console.error("Error fetching data from API:", error);
    }
    return secretaryExists;
  };

  const handleCommentsMobile = () => {
    const isMobile = window.innerWidth <= 767;
    if (isMobile) {
      setMobileCommentsDialog(true);
    }
  };

  const handleCommentsEditMobile = (comment, index) => {
    const isMobile = window.innerWidth <= 767;
    if (isMobile) {
      setCurrentEditComment(comment); // Set the current comment to be edited
      setEditIndex(index);
      setMobileCommentsEditDialog(true);
      setDialogEdit(comment);
    }
  };

  const handleCloseCommentsDialog = () => {
    setMobileCommentsDialog(false);
    setCurrentEditComment("");
    setMobileCommentsEditDialog(false);
  };

  const handleOpenCommentDialog = () => {
    setMobileCommentsDialog(true);
  };

  const handelGenralComments = async (e) => {
    // Added passcode redirect page --> Kavya (24/07) feedback -- 427
    if (verifyPasscode === "true") {
      const { name, value } = e.target;
      // const allowedCharsRegex = /^[a-zA-Z0-9.,:;?!_@#₹ \t\n&"'\/\-\$%]*$/;

      if (name === "PageNo") {
        const trimmedValue = value.replace(/\s/g, "");
        // if (allowedCharsRegex.test(trimmedValue)) {
        if (trimmedValue.length <= 5) {
          SetGeneralCmtInfoObj({
            ...generalcmtInfoobj,
            [name]: trimmedValue,
          });
        } else {
          setErrorMessageEdit("");
          setUserNotifyInfo("Cannot add more than 5 characters.");
          setUserComboValidationDialog(true);
        }
        // } else {
        //   setErrorMessageEdit("Special characters are not allowed.");
        // }
      } else {
        // if (allowedCharsRegex.test(value)) {
        SetGeneralCmtInfoObj({
          ...generalcmtInfoobj,
          [name]: value,
        });
        setErrorMessageEdit("");
        // } else {
        //   setErrorMessageEdit("Special characters are not allowed.");
        // }
      }
    }
    // Added passcode redirect page --> feedback 427 Kavya (24/07)
    else if (verifyPasscode === "false") {
      setValidPasscode(true);
      // Changed the content --> Kavya (25-07)
      setValidmsg(
        "Passcode is not set. Please create passcode to proceed further."
      );
    }
  };
  // on change handler for edit fields
  const handelGenralCommentsEdit = (e) => {
    const { name, value, id } = e.target;
    // Added special chars validation for PageNo, DocRef, and Comments fields in edit comments --> (08-10)
    // const allowedCharsRegex = /^[a-zA-Z0-9.,:;?!_@#₹ \t\n&"'\/\-\$%]*$/;

    const editGrid = generalcmtforAllCmt.map((obj, index) => {
      if (index === Number(id) && name === "PageNo") {
        const trimmedValue = value.replace(/\s/g, "");

        // if (allowedCharsRegex.test(trimmedValue)) {
        if (trimmedValue.length <= 5) {
          return {
            ...obj,
            [name]: trimmedValue,
          };
        } else {
          setErrorMessageEdit("");
          setUserNotifyInfo("Cannot add more than 5 characters.");
          setUserComboValidationDialog(true);
          return obj;
        }
        // } else {
        //   setErrorMessageEdit("Special characters are not allowed.");
        //   return obj;
        // }
      } else if (index === Number(id) && (name === "DocRef" || name === "Comments" || name === "PageNo")) {
        setErrorMessageEdit("");
        // if (allowedCharsRegex.test(value)) {
        return {
          ...obj,
          [name]: value,
        };
        // } else {
        //   setErrorMessageEdit("Special characters are not allowed.");
        //   return obj;
        // }
      } else {
        return obj;
      }
    });

    SetGeneralcmtforAllCmt(editGrid);
  };

  // save edited comments
  // const handleGeneraleditedcommentsSave = (ind) => {
  //   const cmtsObj = [];
  //   setErrorMessageEdit("");
  //   generalcmtforAllCmt.map((obj, index) => {
  //     if (index === ind) {
  //       cmtsObj.push({
  //         DocRef: obj.DocRef.trim() !== "" ? obj.DocRef : "NA",
  //         Comments: obj.Comments.trim() !== "" ? obj.Comments : "NA",
  //         PageNo: obj.PageNo.trim() !== "" ? obj.PageNo : "NA",
  //       });
  //     }
  //   });

  //   if ((cmtsObj[0]?.Comments).trim() !== "") {
  //     const editGrid = generalcmtforAllCmt.map((obj, index) => {
  //       if (index === ind) {
  //         return {
  //           ...obj,
  //           isEdit: false,
  //           PageNo: cmtsObj[0].PageNo, // Ensure "NA" is set if needed
  //           DocRef: cmtsObj[0].DocRef, // Ensure "NA" is set if needed
  //         };
  //       }
  //       return obj;
  //     });
  //     SetGeneralcmtforAllCmt(editGrid);
  //     setMobileCommentsEditDialog(false);
  //   } else {
  //     setUserNotifyInfo("Please fill in Comments fields and then click Save.");
  //     setUserComboValidationDialog(true);
  //   }
  // };

  // const handleGeneraleditedcommentsSave = async (ind) => {
  //   const cmtsObj = [];
  //   setErrorMessageEdit("");

  //   generalcmtforAllCmt.map((obj, index) => {
  //     if (index === ind) {
  //       cmtsObj.push({
  //         DocRef: obj.DocRef.trim() !== "" ? obj.DocRef : "NA",
  //         Comments: obj.Comments.trim() !== "" ? obj.Comments : "NA",
  //         PageNo: obj.PageNo.trim() !== "" ? obj.PageNo : "NA",
  //         noteApproverCommentID: obj.noteApproverCommentID || 0,
  //       });
  //     }
  //   });

  //   if ((cmtsObj[0]?.Comments || "").trim() !== "") {
  //     const editGrid = generalcmtforAllCmt.map((obj, index) => {
  //       if (index === ind) {
  //         return {
  //           ...obj,
  //           isEdit: false,
  //           PageNo: cmtsObj[0].PageNo,
  //           DocRef: cmtsObj[0].DocRef,
  //         };
  //       }
  //       return obj;
  //     });
  //     SetGeneralcmtforAllCmt(editGrid);
  //     setMobileCommentsEditDialog(false);

  //     try {
  //       const accessToken = await getAccessToken({ ...loginRequest, account }, instance);

  //       const approverObj = noteData.noteApproversDTO.find(
  //         (item) => item.approverEmail?.toLowerCase() === accounts[0].username.toLowerCase()
  //       );
  //       const noteApproverId = approverObj?.noteApproverId;

  //       if (!noteApproverId) {
  //         setErrorMessageEdit("Unable to find your approver ID.");
  //         return;
  //       }

  //       const payload = {
  //         NoteApproverId: noteApproverId,
  //         NoteApproverCommentsId: cmtsObj[0].noteApproverCommentID,
  //         PageNo: cmtsObj[0].PageNo,
  //         DocRef: cmtsObj[0].DocRef,
  //         CommentsLog: cmtsObj[0].Comments,
  //         CreatedBy: accounts[0].username,
  //       };

  //       await fetch(`${API_BASE_URL}${API_ENDPOINTS.SaveApproverComment}`, {
  //         method: "POST",
  //         headers: {
  //           ...API_COMMON_HEADERS,
  //           Authorization: `Bearer ${accessToken}`,
  //           "Content-Type": "application/json",
  //         },
  //         body: JSON.stringify(payload),
  //       });
  //     } catch (err) {
  //       console.error("Error saving edited comment:", err);
  //       setErrorMessageEdit("Failed to save comment. Please try again.");
  //     }
  //   } else {
  //     setUserNotifyInfo("Please fill in Comments fields and then click Save.");
  //     setUserComboValidationDialog(true);
  //   }
  // };

  const handleGeneraleditedcommentsSave = async (ind) => {
    const cmtsObj = [];
    setErrorMessageEdit("");

    generalcmtforAllCmt.map((obj, index) => {
      if (index === ind) {
        // cmtsObj.push({
        //   DocRef: obj.DocRef.trim() !== "" ? obj.DocRef : "NA",
        //   Comments: obj.Comments.trim() !== "" ? obj.Comments : "NA",
        //   PageNo: obj.PageNo.trim() !== "" ? obj.PageNo : "NA",
        //   noteApproverCommentID: obj.noteApproverCommentID,
        //   noteReferrerCommentID: referedID, 
        // });
        cmtsObj.push({
          DocRef: obj.DocRef.trim(),
          Comments: obj.Comments.trim(),
          PageNo: obj.PageNo.trim(),
          noteApproverCommentID: obj.noteApproverCommentID,
          noteReferrerCommentID: referedID,
        });

      }
    });

    if ((cmtsObj[0]?.Comments || "").trim() !== "") {
      const editGrid = generalcmtforAllCmt.map((obj, index) => {
        if (index === ind) {
          return {
            ...obj,
            isEdit: false,
            PageNo: cmtsObj[0].PageNo,
            DocRef: cmtsObj[0].DocRef,
          };
        }
        return obj;
      });

      SetGeneralcmtforAllCmt(editGrid);
      setMobileCommentsEditDialog(false);

      try {
        const accessToken = await getAccessToken({ ...loginRequest, account }, instance);

        const isReferBackMode =
          noteData?.currentActioner === accounts[0].username &&
          isreferredUser &&
          noteData.status === 9 &&
          noteData !== null;

        if (isReferBackMode) {
          // 🔁 Save edited refer back comment
          const referrerObj = noteData.noteReferrerDTO.find(
            (ref) => ref.referrerEmail?.toLowerCase() === accounts[0].username.toLowerCase()
          );
          const noteReferrerId = referrerObj?.noteReferrerId;

          if (!noteReferrerId) {
            setErrorMessageEdit("Unable to find your referrer ID.");
            return;
          }

          const referBackPayload = {
            NoteReferrerCommentId: referedID, // Use existing ID
            NoteReferrerId: noteReferrerId,
            PageNo: cmtsObj[0].PageNo,
            DocRef: cmtsObj[0].DocRef,
            CommentsLog: cmtsObj[0].Comments,
            CreatedBy: accounts[0].username,
          };

          await fetch(`${API_BASE_URL}/api/NoteReferre/SaveReffererComment`, {
            method: "POST",
            headers: {
              ...API_COMMON_HEADERS,
              Authorization: `Bearer ${accessToken}`,
              "Content-Type": "application/json",
            },
            body: JSON.stringify(referBackPayload),
          });
        } else {
          // ✏️ Save edited approver comment
          const approverObj = noteData.noteApproversDTO.find(
            (item) => item.approverEmail?.toLowerCase() === accounts[0].username.toLowerCase()
          );
          const noteApproverId = approverObj?.noteApproverId;

          if (!noteApproverId) {
            setErrorMessageEdit("Unable to find your approver ID.");
            return;
          }

          const approverPayload = {
            NoteApproverId: noteApproverId,
            NoteApproverCommentsId: cmtsObj[0].noteApproverCommentID,
            PageNo: cmtsObj[0].PageNo,
            DocRef: cmtsObj[0].DocRef,
            CommentsLog: cmtsObj[0].Comments,
            CreatedBy: accounts[0].username,
          };

          await fetch(`${API_BASE_URL}${API_ENDPOINTS.SaveApproverComment}`, {
            method: "POST",
            headers: {
              ...API_COMMON_HEADERS,
              Authorization: `Bearer ${accessToken}`,
              "Content-Type": "application/json",
            },
            body: JSON.stringify(approverPayload),
          });
        }
      } catch (err) {
        console.error("Error saving edited comment:", err);
        setErrorMessageEdit("Failed to save comment. Please try again.");
      }
    } else {
      setUserNotifyInfo("Please fill in Comments fields and then click Save.");
      setUserComboValidationDialog(true);
    }
  };


  // const handleGeneraleditedcommentsSave = (ind) => {
  //   const cmtsObj = [];
  //   setErrorMessageEdit("");
  //   generalcmtforAllCmt.map((obj, index) => {
  //     if (index === ind) {
  //       cmtsObj.push({
  //         DocRef: obj.DocRef.trim() !== "" ? obj.DocRef : "NA",
  //         Comments: obj.Comments.trim() !== "" ? obj.Comments : "NA",
  //         PageNo: obj.PageNo.trim() !== "" ? obj.PageNo : "NA",
  //       });
  //     }
  //   });

  //   if ((cmtsObj[0]?.Comments).trim() !== "") {
  //     const editGrid = generalcmtforAllCmt.map((obj, index) => {
  //       if (index === ind) {
  //         return {
  //           ...obj,
  //           isEdit: false,
  //           PageNo: cmtsObj[0].PageNo, // Ensure "NA" is set if needed
  //           DocRef: cmtsObj[0].DocRef, // Ensure "NA" is set if needed
  //         };
  //       }
  //       return obj;
  //     });
  //     SetGeneralcmtforAllCmt(editGrid);
  //     setMobileCommentsEditDialog(false);
  //   } else {
  //     setUserNotifyInfo("Please fill in Comments fields and then click Save.");
  //     setUserComboValidationDialog(true);
  //   }
  // };

  // enable inputs onclick of edit
  const handleMobileSave = async () => {
    const cmtsObj = [];

    generalcmtforAllCmt.forEach((obj) => {
      if (obj.isEdit) {
        cmtsObj.push({
          DocRef: obj.DocRef.trim(),
          Comments: obj.Comments.trim(),
          PageNo: obj.PageNo.trim(),
          noteApproverCommentID: obj.noteApproverCommentID,
          noteReferrerCommentID: referedID,
        });
      }
    });

    if ((cmtsObj[0]?.Comments || "").trim() !== "") {
      const editGrid = generalcmtforAllCmt.map((obj) => {
        if (obj.isEdit) {
          return {
            ...obj,
            isEdit: false,
            PageNo: cmtsObj[0].PageNo,
            DocRef: cmtsObj[0].DocRef,
            Comments: cmtsObj[0].Comments,
          };
        }
        return obj;
      });

      SetGeneralcmtforAllCmt(editGrid);
      handleCloseCommentsDialog();

      try {
        const accessToken = await getAccessToken({ ...loginRequest, account }, instance);

        const isReferBackMode =
          noteData?.currentActioner === accounts[0].username &&
          isreferredUser &&
          noteData.status === 9 &&
          noteData !== null;

        if (isReferBackMode) {
          const referrerObj = noteData.noteReferrerDTO.find(
            (ref) => ref.referrerEmail?.toLowerCase() === accounts[0].username.toLowerCase()
          );
          const noteReferrerId = referrerObj?.noteReferrerId;

          if (!noteReferrerId) {
            setUserNotifyInfo("Unable to find your referrer ID.");
            setUserComboValidationDialog(true);
            return;
          }

          const referBackPayload = {
            NoteReferrerCommentId: cmtsObj[0].noteReferrerCommentID || 0,
            NoteReferrerId: noteReferrerId,
            PageNo: cmtsObj[0].PageNo,
            DocRef: cmtsObj[0].DocRef,
            CommentsLog: cmtsObj[0].Comments,
            CreatedBy: accounts[0].username,
          };

          await fetch(`${API_BASE_URL}/api/NoteReferre/SaveReffererComment`, {
            method: "POST",
            headers: {
              ...API_COMMON_HEADERS,
              Authorization: `Bearer ${accessToken}`,
              "Content-Type": "application/json",
            },
            body: JSON.stringify(referBackPayload),
          });

        } else {
          const approverObj = noteData.noteApproversDTO.find(
            (item) => item.approverEmail?.toLowerCase() === accounts[0].username.toLowerCase()
          );
          const noteApproverId = approverObj?.noteApproverId;

          if (!noteApproverId) {
            setUserNotifyInfo("Unable to find your approver ID.");
            setUserComboValidationDialog(true);
            return;
          }

          const approverPayload = {
            NoteApproverId: noteApproverId,
            NoteApproverCommentsId: cmtsObj[0].noteApproverCommentID,
            PageNo: cmtsObj[0].PageNo,
            DocRef: cmtsObj[0].DocRef,
            CommentsLog: cmtsObj[0].Comments,
            CreatedBy: accounts[0].username,
          };

          await fetch(`${API_BASE_URL}${API_ENDPOINTS.SaveApproverComment}`, {
            method: "POST",
            headers: {
              ...API_COMMON_HEADERS,
              Authorization: `Bearer ${accessToken}`,
              "Content-Type": "application/json",
            },
            body: JSON.stringify(approverPayload),
          });
        }
      } catch (err) {
        console.error("Error saving mobile comment:", err);
        setUserNotifyInfo("Failed to save comment. Please try again.");
        setUserComboValidationDialog(true);
      }
    } else {
      setUserNotifyInfo("Please fill in the Comments field and then click Save.");
      setUserComboValidationDialog(true);
    }
  };
  const handelingEdit = (ind) => {
    const editGrid = generalcmtforAllCmt.map((obj, index) => {
      if (index === ind) {
        return {
          ...obj,
          isEdit: true,
        };
      } else {
        return {
          ...obj,
          isEdit: false,
        };
      }
    });
    SetGeneralcmtforAllCmt(editGrid);
    setDialogEdit(editGrid);
  };

  // ========== PDF ANNOTATION CHANGE START ==========
  // MODIFIED - Added filtering to exclude PDF annotations from General Comments table
  // Previously: All comments were shown in General Comments table
  // Now: PDF comments are shown only on PDF, not in General Comments table
  useEffect(() => {
    if (!Array.isArray(noteData?.noteApproverCommentsDTO) || !accounts?.[0]?.username) {
      SetGeneralcmtforAllCmt([]);
      return;
    }

    const currentUser = (accounts[0].username || "").toLowerCase();
    const isFinalStatus = [5, 6, 7].includes(Number(noteData?.status || 0));

    if (isFinalStatus) {
      SetGeneralcmtforAllCmt([]);
      return;
    }

    const filteredComments = noteData.noteApproverCommentsDTO
      .filter((cmt) => {
        const createdBy = (cmt?.createdBy || "").toLowerCase();
        const commentText = cmt?.comments || "";

        const isOwnComment = createdBy === currentUser;

        // only normal editable comments, not workflow-generated comment rows
        const isDirectComment =
          cmt?.workflowLogId === null ||
          cmt?.workflowLogId === 0 ||
          cmt?.workflowLogId === undefined;

        const isPdfHighlight = commentText.startsWith("[PDF Highlight]:");
        const isPdfComment = commentText.startsWith("[PDF Comment]:");

        return isOwnComment && isDirectComment && !isPdfHighlight && !isPdfComment;
      })
      .sort((a, b) => new Date(b.createdDate) - new Date(a.createdDate))
      .map((cmt) => ({
        noteApproverCommentID: cmt.noteApproverCommentID || cmt.noteApproverCommentsId || 0,
        PageNo: cmt.pageNumber || "",
        DocRef: cmt.docReferrence || "",
        Comments: cmt.comments || "",
        createdBy: cmt.createdBy || "",
        createdDate: cmt.createdDate || "",
        workflowLogId: cmt.workflowLogId || 0,
        annotationId: cmt.annotationId || null,
        isEdit: false,
      }));

    SetGeneralcmtforAllCmt(filteredComments);
  }, [noteData, accounts]);
  // ========== PDF ANNOTATION CHANGE END ==========




  const handelAddCmts = async (e) => {
    e.preventDefault();
    setErrorMessageEdit("");
    if (isSubmitting) return;
    setIsSubmitting(true);
    try {
      if (verifyPasscode === "true") {
        if (generalcmtInfoobj.Comments !== "") {

          const updatedGeneralCmtInfoObj = {
            ...generalcmtInfoobj,
            DocRef: generalcmtInfoobj.DocRef.trim(),
            PageNo: generalcmtInfoobj.PageNo.trim(),
          };

          try {
            const accessToken = await getAccessToken({ ...loginRequest, account }, instance);

            // Check if we should call the Refer Back API
            const isReferBackMode =
              noteData?.currentActioner === accounts[0].username &&
              isreferredUser &&
              noteData?.status === 9 &&
              noteData !== null;

            let newCommentEntry = {
              ...updatedGeneralCmtInfoObj,
              isEdit: false,
            };

            if (isReferBackMode) {
              // Call the Refer Back API
              const referrerObj = noteData.noteReferrerDTO.find(
                (ref) => ref.referrerEmail?.toLowerCase() === accounts[0].username.toLowerCase()
              );
              const noteReferrerId = referrerObj?.noteReferrerId;

              if (!noteReferrerId) {
                setErrorMessageEdit("Unable to find your referrer ID.");
                return;
              }

              const referBackPayload = {
                NoteReferrerCommentId: 0, // New comment
                NoteReferrerId: noteReferrerId,
                PageNo: updatedGeneralCmtInfoObj.PageNo,
                DocRef: updatedGeneralCmtInfoObj.DocRef,
                CommentsLog: updatedGeneralCmtInfoObj.Comments,
                CreatedBy: accounts[0].username,
              };

              const response = await fetch(`${API_BASE_URL}/api/NoteReferre/SaveReffererComment`, {
                method: "POST",
                headers: {
                  ...API_COMMON_HEADERS,
                  Authorization: `Bearer ${accessToken}`,
                  "Content-Type": "application/json",
                },
                body: JSON.stringify(referBackPayload),
              });

              if (!response.ok) throw new Error("Failed to save comment.");

              const responseJson = await response.json();
              console.log(responseJson, "responseJson");
              setReferedID(responseJson.noteReferrerCommentId);

              const newCommentId = responseJson?.noteReferrerCommentId;

              newCommentEntry = {
                ...newCommentEntry,
                noteReferrerCommentID: newCommentId,
              };
            } else {
              // Call the Approver Comment API
              const currentApprover = noteData.noteApproversDTO.find(
                (approver) => approver.approverEmail?.toLowerCase() === accounts[0].username.toLowerCase()
              );

              const noteApproverId = currentApprover?.noteApproverId;

              if (!noteApproverId) {
                setErrorMessageEdit("Unable to find your approver ID.");
                return;
              }

              const approverPayload = {
                NoteApproverId: noteApproverId,
                NoteApproverCommentsId: 0, // New comment
                PageNo: updatedGeneralCmtInfoObj.PageNo,
                DocRef: updatedGeneralCmtInfoObj.DocRef,
                CommentsLog: updatedGeneralCmtInfoObj.Comments,
                CreatedBy: accounts[0].username,
              };

              const response = await fetch(`${API_BASE_URL}${API_ENDPOINTS.SaveApproverComment}`, {
                method: "POST",
                headers: {
                  ...API_COMMON_HEADERS,
                  Authorization: `Bearer ${accessToken}`,
                  "Content-Type": "application/json",
                },
                body: JSON.stringify(approverPayload),
              });

              if (!response.ok) throw new Error("Failed to save comment.");

              const responseJson = await response.json();
              const newCommentId = responseJson?.noteApproverCommentsId;
              const newReferedCommentId = responseJson?.noteReferrerCommentId;

              newCommentEntry = {
                ...newCommentEntry,
                noteApproverCommentID: newCommentId,
                noteReferrerCommentId: newReferedCommentId
              };
            }

            // Update comment list
            SetGeneralcmtforAllCmt((prev) => [newCommentEntry, ...prev]);

            // Clear input
            SetGeneralCmtInfoObj({
              DocRef: "",
              Comments: "",
              PageNo: "",
              isEdit: false,
            });

            setMobileCommentsDialog(false);
          } catch (err) {
            console.error("Error saving comment:", err);
            setErrorMessageEdit("Failed to save comment. Please try again.");
          }
        } else {
          setUserNotifyInfo("Please fill in the Comments field and then click Add Comments.");
          setUserComboValidationDialog(true);
        }
      } else {
        setValidPasscode(true);
        setValidmsg("Passcode is not set. Please create passcode to proceed further.");
      }
    } finally {
      setIsSubmitting(false);
      setActionInProgress(false);
    }
  };



  const handelDeleteCmt = (ind) => {
    setSelectedDeleteIndex(ind);
    setDeleteNotification(true);
  };


  const handleDeleteConfirmation = async () => {
    const ind = selectedDeleteIndex;
    const commentToDelete = generalcmtforAllCmt[ind];

    const isReferBackMode =
      noteData?.currentActioner === accounts[0].username &&
      isreferredUser &&
      noteData?.status === 9 &&
      noteData !== null;

    console.log("Refer Back Mode:", isReferBackMode);
    console.log("Comment to delete:", commentToDelete);

    let commentId = null;
    let idKey = null;

    if (isReferBackMode) {
      // Attempt to extract noteReferrerCommentId from noteData.noteReferrerCommentsDTO


      const noteReffererCommentsId = noteData?.noteReferrerCommentsDTO?.filter(
        (obj) =>
          obj.comments === commentToDelete.Comments &&
          obj.pageNumber === commentToDelete.PageNo &&
          obj.docReferrence === commentToDelete.DocRef
      )[0]?.noteReferrerCommentId;

      console.log(noteReffererCommentsId, "noteReffererCommentsId");


      commentId = noteReffererCommentsId;
      idKey = "noteReffererCommentsId";

      console.log(commentId, "commentId");


      console.log("Resolved noteReffererCommentsId:", commentId);
    } else {
      commentId = commentToDelete?.noteApproverCommentID;
      idKey = "noteApproverCommentsId";

      console.log("Resolved noteApproverCommentID:", commentId);
    }

    if (!commentId) {
      console.warn(`No ${idKey} found. Skipping API call.`);
      SetGeneralcmtforAllCmt(
        generalcmtforAllCmt.filter((_, index) => index !== ind)
      );
      setDeleteNotification(false);
      return;
    }

    try {
      const accessToken = await getAccessToken(
        { ...loginRequest, account },
        instance
      );

      const queryParams = new URLSearchParams({
        [idKey]: commentId,
      }).toString();

      const deleteUrl = isReferBackMode
        ? `${API_BASE_URL}/api/NoteReferre/DeleteReffererComment?${queryParams}`
        : `${API_BASE_URL}${API_ENDPOINTS.deleteComment}?${queryParams}`;

      console.log("Calling DELETE API:", deleteUrl);

      const deleteResponse = await fetch(deleteUrl, {
        method: "POST",
        headers: {
          ...API_COMMON_HEADERS,
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
      });

      if (!deleteResponse.ok) {
        throw new Error("Delete API failed.");
      }

      // Update UI after successful deletion
      SetGeneralcmtforAllCmt(
        generalcmtforAllCmt.filter((_, index) => index !== ind)
      );
      setDeleteNotification(false);
      setSelectedDeleteIndex(null);
    } catch (err) {
      console.error("Error deleting comment:", err);
      setUserNotifyInfo("Failed to delete comment. Please try again.");
      setUserComboValidationDialog(true);
      setDeleteNotification(false);
    }
  };


  // add referr comments
  const handelReferGenralComments = (e) => {
    const { name, value } = e.target;
    setReferCommentsObj({
      ...generalcmtInfoobj,
      [name]: value,
      DocRef: "",
      PageNo: "",
      isEdit: true,
    });
  };

  const handleComboChangeAssignees = async (event) => {
    // Added passcode redirect page --> feedback 427 Kavya (23/07)
    if (verifyPasscode === "true") {
      setSelectedAssignee(event.value);
    } else {
      setValidPasscode(true);
      // Changed the content --> Kavya (25-07)
      setValidmsg(
        "Passcode is not set. Please create passcode to proceed further."
      );
    }
  };

  // add Assignees ATR
  // const handleAddRowAssignees = async () => {
  //   if (verifyPasscode === "true") {
  //     if (!selectedAssignee) {
  //       setUserNotifyInfo("Please select the Assignee then click on Add.");
  //       setUserComboValidationDialog(true);
  //       return; // Prevent further execution
  //     }

  //     // Check if the selected assignee already exists in selectedAssigneeUsersinfo or noteData.atrAssigneesDTO
  //     const objExistInSelected = selectedAssigneeUsersinfo?.some(
  //       (obj) => obj.atrAssignerEmailName === selectedAssignee
  //     );
  //     const objExistInTable = noteData?.atrAssigneesDTO?.some(
  //       (obj) => obj.atrAssignerEmailName === selectedAssignee
  //     );

  //     if (objExistInSelected || objExistInTable) {
  //       setUserNotifyInfo(
  //         "The selected assignee already exist. Kindly choose another assignee."
  //       );
  //       setUserComboValidationDialog(true);
  //       setSelectedAssignee(""); // Clear selected assignee
  //       return; // Prevent further execution
  //     }

  //     // If the assignee does not exist, add it
  //     const atrAssigneesArray = [];
  //     const currentATRCreatorObj = aTRCreaters?.find(
  //       (obj) => obj.atrCreatorEmail === currentActionerEmail
  //     );
  //     atrAssigneesArray.push({
  //       ATRCreatorEmail: currentATRCreatorObj?.atrCreatorEmail || "",
  //       atrAssignerEmail: assigneeUsers?.find(
  //         (obj) => obj.atrAssignerEmailName === selectedAssignee
  //       )?.atrAssignerEmail || internalUsers?.find((obj) => obj.approverEmailName === selectedAssignee)?.approverEmail || "",
  //       noteRequesterComments: `${generalcmtforAllCmt
  //         ?.map((obj) => `${obj.PageNo} ${obj.DocRef} ${obj.Comments}`)
  //         .join(", ")}`,
  //       createdDate: new Date(),
  //       createdBy: accounts[0].username,
  //       modifiedDate: new Date(),
  //       modifiedBy: accounts[0].username,
  //       atrAssignerEmailName: selectedAssignee,
  //     });

  //     // Update the state with the new assignee
  //     setSelectedAssignee("");
  //     setSelectedAssigneeUsersinfo([
  //       ...selectedAssigneeUsersinfo,
  //       ...atrAssigneesArray,
  //     ]);
  //   } else {
  //     setValidPasscode(true);
  //     setValidmsg("Passcode is not set. Please create a passcode to proceed further.");
  //   }
  // };

  const handleAddRowAssignees = async () => {
    if (verifyPasscode !== "true") {
      setValidPasscode(true);
      setValidmsg("Passcode is not set. Please create a passcode to proceed further.");
      return;
    }

    // Comments are always mandatory for assigning ATR (regardless of final approver)
    const hasValidComments = generalcmtforAllCmt?.some(
      (obj) => obj.Comments?.trim() !== ""
    );
    if (!hasValidComments) {
      setUserNotifyInfo("Comments are Mandatory for assigning ATR.");
      setUserComboValidationDialog(true);
      return;
    }

    // Check if assignee is selected
    if (!selectedAssignee) {
      setUserNotifyInfo("Please select the Assignee then click on Add.");
      setUserComboValidationDialog(true);
      return;
    }

    // Check if the assignee already exists
    const objExistInSelected = selectedAssigneeUsersinfo?.some(
      (obj) => obj.atrAssignerEmailName === selectedAssignee
    );
    const objExistInTable = noteData?.atrAssigneesDTO?.some(
      (obj) => obj.atrAssignerEmailName === selectedAssignee
    );

    if (objExistInSelected || objExistInTable) {
      setUserNotifyInfo("The selected assignee already exist. Kindly choose another assignee.");
      setUserComboValidationDialog(true);
      setSelectedAssignee("");
      return;
    }

    // Add the new assignee
    const atrAssigneesArray = [];
    const currentATRCreatorObj = aTRCreaters?.find(
      (obj) => obj.atrCreatorEmail === currentActionerEmail
    );

    atrAssigneesArray.push({
      ATRCreatorEmail: currentATRCreatorObj?.atrCreatorEmail || "",
      atrAssignerEmail:
        assigneeUsers?.find((obj) => obj.atrAssignerEmailName === selectedAssignee)?.atrAssignerEmail ||
        internalUsers?.find((obj) => obj.approverEmailName === selectedAssignee)?.approverEmail ||
        "",
      noteRequesterComments: `${generalcmtforAllCmt
        ?.map((obj) => `${obj.PageNo} ${obj.DocRef} ${obj.Comments}`)
        .join(", ")}`,
      createdDate: new Date(),
      createdBy: accounts[0].username,
      modifiedDate: new Date(),
      modifiedBy: accounts[0].username,
      atrAssignerEmailName: selectedAssignee,
    });

    setSelectedAssignee("");
    setSelectedAssigneeUsersinfo([
      ...selectedAssigneeUsersinfo,
      ...atrAssigneesArray,
    ]);
  };

  // date converion
  const getdateconversion = (date) => {
    const convertedDate = new DateObject(new Date(date)).format(
      "DD-MMM-YYYY hh:mm A"
    );
    return convertedDate;
  };

  // get all supporting sizes
  const getAllSupportingDocSize = () => {
    let totalFileSize = 0;
    const inValidFileNames = Object.keys(supportDocWarning).filter(
      (fileName) => supportDocWarning[fileName]?.isValid === false
    );
    // filter valid files only if()
    if (supportDocfilesInfo.length > 0) {
      const vaildMultiplefile = supportDocfilesInfo.filter((obj) => {
        if (!inValidFileNames.includes(obj.supportingDocumentFileName)) {
          return obj;
        }
      });
      vaildMultiplefile.map(
        (obj) =>
          (totalFileSize = totalFileSize + obj.supportingDocumentPathLength)
      );
    }
    if (noteData?.noteSupportingDocumentsDTO?.length > 0) {
      noteData?.noteSupportingDocumentsDTO.map(
        (obj) =>
          (totalFileSize = totalFileSize + obj.supportingDocumentPathLength)
      );
    }
    return totalFileSize;
  };

  // onchange passcode
  const handlePasscodeChange = (e) => {
    const { value } = e.target;
    // Regular expression to allow only alphanumeric characters
    const alphanumericRegex = /^[a-zA-Z0-9]*$/;

    if (alphanumericRegex.test(value)) {
      setPasscode(value);
    }
  };

  // Verify OTP
  const passcodeVerificationFunction = async () => {
    try {
      const accessToken = await getAccessToken(
        { ...loginRequest, account },
        instance
      );

      // Replace with your own secret key
      const secretKey = "SHA256";

      // Compute HMAC-SHA256 hash of the passcode with the secret key
      const hashedPasscode = CryptoJS.SHA256(passcode, secretKey).toString(
        CryptoJS.enc.Hex
      );

      const params = {
        Passcode: hashedPasscode,
        UserMail: accounts[0].username,
      };

      const response = await fetch(
        `${API_BASE_URL}${API_ENDPOINTS.eNote_VerifyPasscode}`,
        {
          method: "POST",
          body: JSON.stringify(params),
          headers: {
            ...API_COMMON_HEADERS,
            Authorization: `Bearer ${accessToken}`,
            "Content-Type": "application/json",
          },
        }
      );

      const data = await response.text();

      if (data === "Passcode verification successful") {
        // Handle success scenario
        setPasscodeVerification(false);

        // Define confirmDialogSettings (assuming it's defined elsewhere)
        const confirmDialogSettings = {
          Approve: {
            Confirmtext: `Are you sure you want to ${isApprover() ? "note" : "approve"} this request?`,
            Description: `Please check the details filled along with attachment and click on Confirm button to ${isApprover() ? "note" : "approve"} request.`,
          },
          Reject: {
            Confirmtext: "Are you sure you want to reject this request?",
            Description:
              "Please check the details filled along with attachment and click on Confirm button to reject request.",
          },
          "Refer Back": {
            Confirmtext: "Are you sure you want to Refer Back this request?",
            Description:
              "Please check the details filled along with attachment and click on Confirm button to refer back request.",
          },
          Return: {
            Confirmtext: "Are you sure you want to return this request?",
            Description:
              "Please check the details filled along with attachment and click on Confirm button to return request.",
          },
        };

        // Perform actions based on actionBtn
        if (actionBtn in confirmDialogSettings) {
          setConfirmDailogObj(confirmDialogSettings[actionBtn]);
          toggleDailogForConfirmation();
        }

        if (actionBtn === "Call Back") {
          handlecallbackAPICal();
        }

        if (actionBtn === "ChangeApprover") {
          hnadleChangeApproverAPICall();
        }

        if (actionBtn === "Refer") {
          try {
            await Promise.all([approverStatusChange()]);
            setdialogOpen(false);
            setDialogForRefereeAndChangeApprover(false);
          } catch (error) {
            console.error("Error:", error);
          }
        }

        // Return confirmDialogSettings on success
        return confirmDialogSettings;
      } else {
        setError("Incorrect passcode entered please try again.");
      }
    } catch (err) {
      console.error(err);
      // Handle error state or throw it further
      throw err;
    }
  };

  // verify passcode function
  const VerifyUserPasscode = async () => {
    let userPasscodeExist = "false";
    try {
      const accessToken = await getAccessToken(
        { ...loginRequest, account },
        instance
      );
      const userPasscodeResponse = await fetch(
        `${API_BASE_URL}${API_ENDPOINTS.eNote_VerifyUserPasscode(
          accounts[0].username
        )}`,
        {
          method: "POST",
          headers: {
            ...API_COMMON_HEADERS,
            Authorization: `Bearer ${accessToken}`,
            "Content-Type": "application/json",
          },
        }
      );

      const userData = await userPasscodeResponse.text();
      if (userData) {
        userPasscodeExist = userData;
        setVerifypasscode(userData);
      }
    } catch (err) {
      console.error(err);
      throw err;
    }
    return userPasscodeExist;
  };

  // passcode close
  const handlepasscodeClose = () => {
    setPasscodeVerification(false);
    setActionInProgress(false);
    setError(false);
  };

  // all  action btn start here - on Approver btn click
  // const approvalFunction = async () => {
  //   let maxFileSize = 26214400;
  //   let totalFileSize = getAllSupportingDocSize();

  //   if (
  //     Object.keys(supportDocWarning).length > 0 ||
  //     totalFileSize > maxFileSize
  //   ) {
  //     setActionBtn("Approve");
  //     if (totalFileSize > maxFileSize) {
  //       setUserNotifyInfo(
  //         "Cumulative size of all the supporting documents should not be exceeded 25 MB."
  //       );
  //     } else if (Object.keys(supportDocWarning).length > 0) {
  //       setUserNotifyInfo(
  //         "Please check and remove if there are any invalid files attached."
  //       );
  //     }
  //     setUserComboValidationDialog(true);
  //   } else {
  //     const isAtrRequired = noteData?.noteApproversDTO?.some(
  //       (obj) =>
  //         obj.approverEmail === accounts[0].username &&
  //         obj.approverEmail === noteData?.currentActioner &&
  //         obj.approverType === 2 &&
  //         obj.approverStatus === 2
  //     );

  //     if (isAtrRequired && generalcmtforAllCmt?.length === 0) {
  //       setActionBtn("Approve");
  //       setUserNotifyInfo("Please fill in comments then click on Approve.");
  //       setUserComboValidationDialog(true);
  //     } else
  //       // Added passcode redirect page --> Kavya (23/07) feedback -- 427
  //       if (verifyPasscode === "true") {
  //         setPasscode("");
  //         setPasscodeVerification(true);
  //         setActionBtn("Approve");
  //       } else {
  //         setValidPasscode(true);
  //         // Changed the content --> Kavya (25-07)
  //         setValidmsg(
  //           "Passcode is not set. Please create passcode to proceed further."
  //         );
  //         setActionBtn("Approve");
  //       }
  //   }
  // };

  // const approvalFunction = async () => {
  //   let maxFileSize = 26214400;
  //   let totalFileSize = getAllSupportingDocSize();

  //   if (
  //     Object.keys(supportDocWarning).length > 0 ||
  //     totalFileSize > maxFileSize
  //   ) {
  //     setActionBtn("Approve");
  //     if (totalFileSize > maxFileSize) {
  //       setUserNotifyInfo(
  //         "Cumulative size of all the supporting documents should not be exceeded 25 MB."
  //       );
  //     } else {
  //       setUserNotifyInfo(
  //         "Please check and remove if there are any invalid files attached."
  //       );
  //     }
  //     setUserComboValidationDialog(true);
  //     return;
  //   }

  //   const currentUserEmail = accounts[0].username;
  //   const approvers = noteData?.noteApproversDTO || [];

  //   const currentApprover = approvers.find(
  //     (a) =>
  //       a.approverEmail === currentUserEmail &&
  //       a.approverEmail === noteData?.currentActioner &&
  //       a.approverType === 2 &&
  //       a.approverStatus === 2
  //   );

  //   const maxApproverOrder = Math.max(...approvers.map(a => a.approverOrder || 0));
  //   const isFinalApprover = currentApprover?.approverOrder === maxApproverOrder;

  //   if (!isFinalApprover && generalcmtforAllCmt?.length === 0) {
  //     setActionBtn("Approve");
  //     setUserNotifyInfo("Please fill in comments then click on Approve.");
  //     setUserComboValidationDialog(true);
  //     return;
  //   }

  //   if (verifyPasscode === "true") {
  //     setPasscode("");
  //     setPasscodeVerification(true);
  //     setActionBtn("Approve");
  //   } else {
  //     setValidPasscode(true);
  //     setValidmsg(
  //       "Passcode is not set. Please create passcode to proceed further."
  //     );
  //     setActionBtn("Approve");
  //   }
  // };
  const getCurrentUserEmail = () =>
    (accounts?.[0]?.username || "").toLowerCase();

  const isPdfSystemComment = (text = "") =>
    text.startsWith("[PDF Comment]:") || text.startsWith("[PDF Highlight]:");

  const getCurrentUserPdfTextComments = () => {
    const currentUserEmail = getCurrentUserEmail();

    const comments =
      pdfAnnotatorRef.current?.getComments?.() ||
      pdfAnnotatorRef.current?.getAnnotations?.()?.filter(
        (a) => (a?.type || "").toLowerCase() === "comment"
      ) ||
      [];

    return comments.filter((comment) => {
      const commentUser = (
        comment?.user ||
        comment?.createdBy ||
        comment?.approverEmail ||
        ""
      ).toLowerCase();

      return (
        commentUser === currentUserEmail &&
        (comment?.text || "").trim() !== ""
      );
    });
  };

  const getCurrentUserGeneralComments = (includeReferDialogComment = false) => {
    const currentUserEmail = getCurrentUserEmail();

    const existingGeneralComments = Array.isArray(generalcmtforAllCmt)
      ? generalcmtforAllCmt.filter((obj) => {
        const commentText = (obj?.Comments || "").trim();

        if (!commentText || isPdfSystemComment(commentText)) {
          return false;
        }

        const createdBy = (
          obj?.createdBy ||
          obj?.CreatedBy ||
          obj?.approverEmail ||
          obj?.ApproverEmail ||
          ""
        ).toLowerCase();

        // Draft comments may not have createdBy, so treat blank createdBy as current user's draft
        return !createdBy || createdBy === currentUserEmail;
      })
      : [];

    if (
      includeReferDialogComment &&
      (referCommentsObj?.Comments || "").trim() !== ""
    ) {
      existingGeneralComments.push({
        ...referCommentsObj,
        createdBy: accounts?.[0]?.username,
      });
    }

    return existingGeneralComments;
  };

  const hasCurrentUserComments = (includeReferDialogComment = false) => {
    const currentUserGeneralComments =
      getCurrentUserGeneralComments(includeReferDialogComment);

    const currentUserPdfComments = getCurrentUserPdfTextComments();

    return (
      currentUserGeneralComments.length > 0 ||
      currentUserPdfComments.length > 0
    );
  };


  const approvalFunction = async () => {
    setActionInProgress(true);
    let maxFileSize = 26214400;
    let totalFileSize = getAllSupportingDocSize();

    if (
      Object.keys(supportDocWarning).length > 0 ||
      totalFileSize > maxFileSize
    ) {
      setActionBtn("Approve");
      if (totalFileSize > maxFileSize) {
        setUserNotifyInfo(
          "Cumulative size of all the supporting documents should not be exceeded 25 MB."
        );
      } else {
        setUserNotifyInfo(
          "Please check and remove if there are any invalid files attached."
        );
      }
      setUserComboValidationDialog(true);
      return;
    }

    // const currentUserEmail = accounts[0].username;
    // const approvers = noteData?.noteApproversDTO || [];

    // const currentApprover = approvers.find(
    //   (a) =>
    //     a.approverEmail === currentUserEmail &&
    //     a.approverEmail === noteData?.currentActioner &&
    //     a.approverStatus === 2
    // );

    // const maxApproverOrder = Math.max(...approvers.map(a => a.approverOrder || 0));
    // const isFinalApprover = currentApprover?.approverOrder === maxApproverOrder;

    // // Only show the comment validation if current actioner is Approver (approverType === 2)
    // if (
    //   currentApprover?.approverType === 2 &&
    //   !isFinalApprover &&
    //   generalcmtforAllCmt?.length === 0
    // ) {
    //   setActionBtn("Approve");
    //   setUserNotifyInfo("Please fill in comments then click on Approve.");
    //   setUserComboValidationDialog(true);
    //   return;
    // }

    if (verifyPasscode === "true") {
      setPasscode("");
      setPasscodeVerification(true);
      setActionBtn("Approve");
    } else {
      setValidPasscode(true);
      setValidmsg(
        "Passcode is not set. Please create passcode to proceed further."
      );
      setActionBtn("Approve");
    }
  };



  const rejectFunction = async () => {
    setActionInProgress(true);

    if (verifyPasscode !== "true") {
      setValidPasscode(true);
      setValidmsg("Passcode is not set. Please create passcode to proceed further.");
      setActionInProgress(false);
      return;
    }

    let maxFileSize = 26214400;
    let totalFileSize = getAllSupportingDocSize();

    if (Object.keys(supportDocWarning).length > 0 || totalFileSize > maxFileSize) {
      setUserNotifyInfo(
        totalFileSize > maxFileSize
          ? "Cumulative size of all the supporting documents should not be exceeded 25 MB."
          : "Please check and remove if there are any invalid files attached."
      );
      setUserComboValidationDialog(true);
      setActionInProgress(false);
      return;
    }

    if (!hasCurrentUserComments(false)) {
      setUserNotifyInfo(
        "Please add either a general comment or a PDF text comment before rejecting."
      );
      setUserComboValidationDialog(true);
      setActionInProgress(false);
      return;
    }

    setPasscode("");
    setPasscodeVerification(true);
    setActionBtn("Reject");
  };

  const referFunction = async () => {
    setActionInProgress(true);

    if (verifyPasscode !== "true") {
      setValidPasscode(true);
      setValidmsg("Passcode is not set. Please create passcode to proceed further.");
      setActionInProgress(false);
      return;
    }

    setActionBtn("Refer");
    toggleSelectingUser();
  };

  const referBackfunction = async () => {
    setActionInProgress(true);

    let maxFileSize = 26214400;
    let totalFileSize = getAllSupportingDocSize();

    // File validation
    if (Object.keys(supportDocWarning).length > 0 || totalFileSize > maxFileSize) {
      if (totalFileSize > maxFileSize) {
        setUserNotifyInfo("Cumulative size of all the supporting documents should not be exceeded 25 MB.");
      } else if (Object.keys(supportDocWarning).length > 0) {
        setUserNotifyInfo("Please check and remove if there are any invalid files attached.");
      }
      setUserComboValidationDialog(true);
      setActionInProgress(false);
      return;
    }

    // Check passcode
    if (verifyPasscode !== "true") {
      setValidPasscode(true);
      setValidmsg("Passcode is not set. Please create passcode to proceed further.");
      setActionInProgress(false);
      return;
    }

    // CHECK IF USER HAS ANY COMMENTS (GENERAL OR PDF)
    const userHasComments = hasCurrentUserComments();
    console.log("User has comments (Refer Back):", userHasComments); // Debug log

    if (!userHasComments) {
      setUserNotifyInfo("Please add either a general comment or a PDF text comment before referring back.");
      setUserComboValidationDialog(true);
      setActionInProgress(false);
      return;
    }

    // Proceed with refer back
    setPasscode("");
    setPasscodeVerification(true);
    setActionBtn("Refer Back");
  };

  // on return btn click - UPDATED with validation
  const returnFunction = async () => {
    setActionInProgress(true);

    if (verifyPasscode !== "true") {
      setValidPasscode(true);
      setValidmsg("Passcode is not set. Please create passcode to proceed further.");
      setActionInProgress(false);
      return;
    }

    let maxFileSize = 26214400;
    let totalFileSize = getAllSupportingDocSize();

    if (Object.keys(supportDocWarning).length > 0 || totalFileSize > maxFileSize) {
      setUserNotifyInfo(
        totalFileSize > maxFileSize
          ? "Cumulative size of all the supporting documents should not be exceeded 25 MB."
          : "Please check and remove if there are any invalid files attached."
      );
      setUserComboValidationDialog(true);
      setActionInProgress(false);
      return;
    }

    if (!hasCurrentUserComments(false)) {
      setUserNotifyInfo(
        "Please add either a general comment or a PDF text comment before returning."
      );
      setUserComboValidationDialog(true);
      setActionInProgress(false);
      return;
    }

    setPasscode("");
    setPasscodeVerification(true);
    setActionBtn("Return");
  };

  //  on  click of gistDoc -submit btn


  const onSubmitForAttcGISTDoc = () => {
    // change 16/04 Gist doc is mandatory

    if (wordandPdfInfo?.wordInfo?.isValid) {
      setActionBtn("AttchGistDoc");
      setConfirmDailogObj({
        Confirmtext: "Are you sure you want to submit this request?",
        Description: "Please click on Confirm button to submit request.",
      });
      toggleDailogForConfirmation();
    } else {
      setActionBtn("AttchGistDoc");
      setUserNotifyInfo("Please select File.");
      setUserComboValidationDialog(true);
    }
  };

  // Function to handle the confirmation
  // const handleConfirmSubmitGistDoc = async () => {
  //   try {
  //     await uploadFiles(fileQueue); // Call uploadFiles after user confirms
  //     console.log("Files uploaded successfully.");
  //   } catch (error) {
  //     console.error("Error during file upload:", error);
  //   }
  // };

  //  on  changeApprover btn click
  const onClickChangeApprover = async () => {
    setActionBtn("ChangeApprover");
    setDialogForRefereeAndChangeApprover(true);
  };

  //on click of confirmation btn it will call the api based on current  action.
  const onCofiormation = async () => {
    if (
      actionBtn === "Approve" ||
      actionBtn === "Return" ||
      actionBtn === "Reject"
    ) {
      await Promise.all([approverStatusChange()])
        .then(() => {
          setdialogOpen(false);
        })
        .catch((error) => {
          console.error("Error:", error);
        });
    }

    if (actionBtn === "AttchGistDoc") {
      addAttachmentsGistDoc();
      setdialogOpen(false);
    }
    if (actionBtn === "Refer Back") {
      onReferredback();
      setdialogOpen(false);
    }
    if (actionBtn === "MarkInfo") {
      addmarkInfotoDataBase();
      setdialogOpen(false);
    }
  };

  // Approval checking if current actioner is approver or reviewer  and is  login user  approval btns will display
  const approverChecking = () => {
    let isCurrentActioner = false;
    if (
      currentActionerEmail === accounts[0].username &&
      (noteData.status === 2 || noteData.status === 3) &&
      noteData !== null &&
      noteData?.noteApproversDTO.some(
        (obj) =>
          obj.approverEmail === currentActionerEmail &&
          obj.approverEmail === accounts[0].username
      )
    ) {
      isCurrentActioner = true;
    }
    return isCurrentActioner;
  };

  // Approver finding
  const isApprover = () => {
    let isCurrentActioner = false;
    if (
      currentActionerEmail === accounts[0].username &&
      noteData?.strNatureofNote === "Information" &&
      (noteData.status === 2 || noteData.status === 3) &&
      noteData !== null &&
      noteData?.noteApproversDTO.some(
        (obj) =>
          obj.approverEmail === currentActionerEmail &&
          obj.approverEmail === accounts[0].username &&
          obj.approverType === 2
      )
    ) {
      isCurrentActioner = true;
    }
    return isCurrentActioner;
  };

  // ATR assignees filtr functionality
  const onFilterChangeForAtrAssigness = (event) => {
    clearTimeout(timeout.current);
    timeout.current = setTimeout(() => {
      setAssigneeUsers(filterAssigneesInfo(event.filter));
    });
  };

  // filter data assignees data
  const filterAssigneesInfo = (filter) => {
    const getAllUser = assigneerUserDetails.filter((item) =>
      item.atrAssignerEmail.toLowerCase().includes(filter.value.toLowerCase())
    );
    return getAllUser;
  };

  //search selected user
  const onFillterALLUser = async (event) => {
    if (event.filter.value.length >= 4) {
      const accessToken = await getAccessToken(
        {
          ...loginRequest,
          account,
        },
        instance
      );
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
          //Change - Adding SR No - 08/04 - RK
          const orgUsers = data.map((x) => {
            return {
              department: x.department === null ? "NA" : x.department,
              displayName: x.displayName === null ? "NA" : x.displayName,
              jobTitle: x.jobTitle === null ? "NA" : x.jobTitle,
              userPrincipalName: x.userPrincipalName,
              srNo: x.srNo === null ? "NA" : x.srNo,
            };
          });
          setOrgEmployees(orgUsers);
        })
        .catch((err) => {
          setOrgEmployees([]);
          console.log(err);
        });
    } else {
      setOrgEmployees([]);
    }
  };

  // dialogs........ all toggle  functions starts here
  // toggle dialog for add refer and change approver
  const toggleSelectingUser = () => {
    setDialogForRefereeAndChangeApprover(true);
  };

  // Success dialog toggle
  const SuccessDialogToggle = () => {
    setDialogSuccess(!dialogSuccess);
  };

  // Failure dialog toggle
  const faildDialogToggle = () => {
    setDialogFailure(true);
  };

  // confirmation dialog toggle
  const toggleDailogForConfirmation = () => {
    setdialogOpen(!dialogOpen);
  };

  //hide dialogs
  // close the callback success dialog
  const handleCloseDialog = () => {
    setCallBackVisible(false);
    setChangeApproverVisible(false);
    setTab("My Pending Notes");
    navigate(redirect);
  };

  //  close closeUseNotify  dialog box
  const oncloseUseNotify = () => {
    const userHasComments = hasCurrentUserComments();

    if (
      actionBtn === "Return" ||
      actionBtn === "Reject" ||
      actionBtn === "Refer Back" ||
      actionBtn === "Approve"
    ) {
      // Only expand General Comments if user has NO comments at all
      if (!userHasComments) {
        setExpanded(["General Comments"]);
        setUserComboValidationDialog(false);
      }
    } else if (actionBtn === "ChangeApprover") {
      setDialogForRefereeAndChangeApprover(true);
      setActionBtn("ChangeApprover");
    }

    setUserComboValidationDialog(false);
    setActionInProgress(false);
  };

  // redirectHomePage  -on click of OK btn in success dialog it will direct to home page
  const redirectHomePage = () => {
    setTab("My Pending Notes");
    navigate(redirect);
    setDialogSuccess(false);
  };

  // Based on approver status rendering diff icons
  const renderApproverIcon = (approverStatus) => {
    let icon = "k-i-clock";
    if (approverStatus === 1) {
      // icon="k-i-rotate-circle"
      icon = "k-i-rotate";
    }
    if (approverStatus === 2) {
      icon = "k-i-clock";
    }
    if (approverStatus === 3) {
      icon = "k-i-check-circle";
    }
    if (approverStatus === 4) {
      icon = "k-i-close-outline k-i-x-outline";
    }
    if (approverStatus === 5) {
      icon = "k-i-undo";
    }
    if (approverStatus === 6) {
      icon = "k-i-arrow-root";
    }
    return icon;
  };

  // Rendering custom Title and Icon for changeApprover/Refer/common dialog box
  const CustomTitleBar = () => {
    return (
      <div
        className="custom-title _render-custon-title"
      >
        <SvgIcon icon={infoCircleIcon} />{" "}
        {actionBtn === "Refer"
          ? " Add Referee"
          : actionBtn === "ChangeApprover"
            ? "Change Approver"
            : "Alert!"}
      </div>
    );
  };

  //Rendering custom Title and Icon for success dialog
  const CustomSuccussTitleBar = () => {
    return (
      <div className="custom-title _render-custon-title">
        <SvgIcon icon={infoCircleIcon} /> Alert!
      </div>
    );
  };

  // Rendering custom Title and Icon for Confirmation dialog.
  const CustomConfirmationTitleBar = () => {
    return (
      <div className="custom-title ">
        <span className="k-icon k-font-icon k-i-borders-show-hide cursor allIconsforPrimary-btn"></span>{" "}
        Confirmation
      </div>
    );
  };

  // API call start here
  const onSubmitforAddReferandChangeApprover = async () => {
    if (actionBtn === "Refer") {
      if (selectedrefer) {
        const isApproverExist = noteData?.noteApproversDTO?.some(
          (obj) => obj.approverEmail === selectedrefer.userPrincipalName
        );

        if (actionBtn === "Refer" && !hasCurrentUserComments(true)) {
          setUserNotifyInfo(
            "Please add either a general comment or a PDF text comment before referring."
          );
          setUserComboValidationDialog(true);
          setActionInProgress(false);
          return;
        }

        if (
          isApproverExist ||
          selectedrefer.userPrincipalName === noteData.createdBy
        ) {
          setUserNotifyInfo(
            "Referrer cannot be the existing approver/reviewer or requester."
          );
          setUserComboValidationDialog(true);
          setdialogOpen(false);
          setDialogForRefereeAndChangeApprover(false);
        } else {
          if (verifyPasscode === "true") {
            setPasscode("");
            setPasscodeVerification(true);
            setDialogForRefereeAndChangeApprover(false);
          } else {
            setValidPasscode(true);
            setDialogForRefereeAndChangeApprover(false);
            setValidmsg(
              "Passcode is not set. Please create passcode to proceed further."
            );
          }
        }
      } else {
        setdialogOpen(false);
        setDialogForRefereeAndChangeApprover(false);
        setUserNotifyInfo("Please select the Referrer then click on Submit.");
        setUserComboValidationDialog(true);
      }
    }

    if (actionBtn === "ChangeApprover") {
      onSubmittingChangeApproverDetails();
    }
  };

  const removeEmptyValues = (obj) => {
    return Object.fromEntries(
      Object.entries(obj).filter(
        ([_, value]) => value !== null && value !== '' && value !== 0
      )
    );
  };
  // ========== PDF ANNOTATION CHANGE START ==========
  // MODIFIED - Added code to extract PDF annotations from PdfAnnotator and include in API payload
  // Also clears draft annotations after successful submission
  const approverStatusChange = async () => {
    debugger;
    setdialogOpen(false);
    setDialogForRefereeAndChangeApprover(false);
    SetIsLoading(true);

    let maxFileSize = 26214400;
    let totalFileSize = 0;

    let selectedAssignees = selectedAssigneeUsersinfo?.map((item) => {
      const { atrAssignerEmailName: _, ...rest } = item;
      return rest;
    });

    // Get invalid file names
    const inValidFileNames = Object.keys(supportDocWarning).filter(
      (fileName) => supportDocWarning[fileName]?.isValid === false
    );

    // Filter valid files only
    const vaildMultiplefile = supportDocfilesInfo.filter((obj) => {
      if (!inValidFileNames.includes(obj.supportingDocumentFileName)) {
        return obj;
      }
      return false;
    });

    vaildMultiplefile.forEach((obj) => {
      totalFileSize = totalFileSize + (obj.supportingDocumentPathLength || 0);
    });

    let fileupdated = vaildMultiplefile?.map((item) => {
      const { supportingDocumentPathLength: _, ...rest } = item;
      return rest;
    });

    const currenArry = noteData.noteApproversDTO?.filter(
      (obj) => obj.approverEmail === accounts[0].username
    );
    const currentObj = currenArry;
    const referredEmail =
      selectedrefer === null ? null : selectedrefer.userPrincipalName;

    const currentUserEmail = (accounts?.[0]?.username || "").toLowerCase();

    // Get general comments - only current user's normal comments
    const generalCmts = Array.isArray(generalcmtforAllCmt)
      ? generalcmtforAllCmt.filter((obj) => {
        const createdBy = (obj?.createdBy || accounts?.[0]?.username || "").toLowerCase();
        const commentText = obj?.Comments || "";

        return (
          createdBy === currentUserEmail &&
          !commentText.startsWith("[PDF Comment]:") &&
          !commentText.startsWith("[PDF Highlight]:")
        );
      })
      : [];

    const ApproverComments = [];

    if (selectedrefer && referCommentsObj?.isEdit) {
      generalCmts.push(referCommentsObj);
    }

    // Add general comments
    generalCmts.forEach((obj) => {
      if (obj?.Comments && obj.Comments.trim() !== "") {
        ApproverComments.push({
          noteApproverCommentID: obj.noteApproverCommentID ?? 0,
          pageNumber: obj.PageNo?.toString() || "NA",
          docReferrence: obj.DocRef || "NA",
          comments: obj.Comments,
          createdDate: new Date().toISOString(),
          createdBy: accounts[0].username,
          modifiedDate: new Date().toISOString(),
          modifiedBy: accounts[0].username,
          noteId: noteData?.noteId,
          approverEmail: accounts[0].username,
          approverType: currentObj.length > 0 ? currentObj[0].approverType : 2,
          approverStatus:
            actionBtn === "Approve"
              ? 3
              : actionBtn === "Reject"
                ? 4
                : actionBtn === "Refer"
                  ? 6
                  : actionBtn === "Return"
                    ? 5
                    : 1,
          strApproverStatus:
            actionBtn === "Approve"
              ? "Approved"
              : actionBtn === "Reject"
                ? "Rejected"
                : actionBtn === "Refer"
                  ? "Referred"
                  : actionBtn === "Return"
                    ? "Returned"
                    : "",
        });
      }
    });



    // 🔥 STEP 1: SAVE HIGHLIGHTS TO BACKEND ON APPROVE

    if (pdfAnnotatorRef.current && includePdfAnnotations) {
      try {
        const allAnnotations = pdfAnnotatorRef.current.getAnnotations?.() || [];

        const currentUserEmail = (accounts?.[0]?.username || "").toLowerCase();

        const unsavedHighlights = allAnnotations.filter((ann) => {
          const type = (ann?.type || "").toLowerCase();
          const user = (ann?.user || "").toLowerCase();

          return (
            type === "highlight" &&
            user === currentUserEmail &&
            !ann.annotationId &&
            !ann.AnnotationId &&
            !ann.serverResponse?.annotationId &&
            Array.isArray(ann.normalizedRects) &&
            ann.normalizedRects.length > 0
          );
        });

        console.log("🔥 Highlights going to backend:", unsavedHighlights);

        for (const highlight of unsavedHighlights) {
          await pdfAnnotatorRef.current.persistAnnotationToServer?.(highlight);
        }

      } catch (err) {
        console.error("❌ Highlight save error:", err);
      }
    }
    // PDF annotations - only current user's annotations
    // In ENoteViewForm.jsx - Replace the pdfCommentEntries section in approverStatusChange function

    let pdfCommentEntries = [];

    if (pdfAnnotatorRef.current && includePdfAnnotations) {
      try {
        // Get all annotations from PdfAnnotator
        const allAnnotations = pdfAnnotatorRef.current.getAnnotations?.() || [];
        const allFormatted = pdfAnnotatorRef.current.getAllFormattedForAPI?.() || [];

        const currentUserEmail = (accounts?.[0]?.username || "").toLowerCase();

        // Map to track annotation IDs from server responses
        const annotationIdMap = new Map();
        allAnnotations.forEach((ann) => {
          const serverId = ann.serverResponse?.annotationId || ann.annotationId || ann.AnnotationId;
          if (serverId) {
            annotationIdMap.set(ann.id, serverId);
          }
        });

        // Filter ONLY comments (not highlights) for current user
        pdfCommentEntries = allAnnotations
          .filter((ann) => {
            const createdBy = (ann?.user || "").toLowerCase();
            return (
              ann.type === "comment" &&
              ann.text?.trim() &&
              createdBy === currentUserEmail
            );
          })
          .map((comment) => {
            const serverAnnotationId = annotationIdMap.get(comment.id) || null;

            // Get line number - prioritize lineNumberRange > lineNumber > fallback
            let lineNo = "NA";
            if (comment.lineNumberRange && comment.lineNumberRange !== "N/A") {
              lineNo = String(comment.lineNumberRange);
            } else if (comment.lineNumber && comment.lineNumber !== "N/A") {
              lineNo = String(comment.lineNumber);
            } else {
              // Try to extract from selected text as fallback
              const extracted = extractLineNumberFromSelectedText?.(comment.selectedText);
              lineNo = extracted ? String(extracted) : "NA";
            }

            return {
              noteApproverCommentID: 0,
              pageNumber: comment.page?.toString() || "NA",
              docReferrence: lineNo,
              comments: `[PDF Comment]: ${comment.text || ""}`,
              createdDate: new Date().toISOString(),
              createdBy: accounts[0].username,
              modifiedDate: new Date().toISOString(),
              modifiedBy: accounts[0].username,
              noteId: noteData?.noteId,
              approverEmail: accounts[0].username,
              approverType: currentObj.length > 0 ? currentObj[0].approverType : 2,
              approverStatus: actionBtn === "Approve" ? 3 : actionBtn === "Reject" ? 4 : actionBtn === "Refer" ? 6 : actionBtn === "Return" ? 5 : 1,
              strApproverStatus: actionBtn === "Approve" ? "Approved" : actionBtn === "Reject" ? "Rejected" : actionBtn === "Refer" ? "Referred" : actionBtn === "Return" ? "Returned" : "",
              annotationId: serverAnnotationId,
              AnnotationId: serverAnnotationId,
              // Store original comment data for debugging
              _originalComment: {
                id: comment.id,
                text: comment.text,
                selectedText: comment.selectedText,
                lineNumber: comment.lineNumber,
                lineNumberRange: comment.lineNumberRange
              }
            };
          })
          .filter(Boolean);

        console.log("✅ PDF comments extracted for NoteApproverComments:", pdfCommentEntries.length);
        pdfCommentEntries.forEach((c, i) => {
          console.log(`  Comment ${i + 1}: Page=${c.pageNumber}, Line=${c.docReferrence}, Text=${c.comments.substring(0, 50)}...`);
        });

      } catch (err) {
        console.error("Error preparing PDF comments:", err);
      }
    }

    const allComments = [...ApproverComments];

    // ✅ Create Set to avoid duplicate comments
    const existingCommentKeys = new Set(
      allComments.map((c) => {
        const keyId = c.annotationId || c.AnnotationId || "";
        return `${c.pageNumber}|${c.docReferrence}|${c.comments}|${keyId}`;
      })
    );
    const extractLineNumberFromSelectedText = (text) => {
      if (!text) return "NA";
      return "NA"; // तुमचा logic नसेल तर default NA
    };

    pdfCommentEntries.forEach((pdfComment) => {
      const keyId = pdfComment.annotationId || pdfComment.pdfAnnotationId || "";
      const key = `${pdfComment.pageNumber}|${pdfComment.docReferrence}|${pdfComment.comments}|${keyId}`;

      if (!existingCommentKeys.has(key)) {
        allComments.push(pdfComment);
        existingCommentKeys.add(key);
      }
    });

    console.log("Total comments being sent:", allComments.length);
    console.log("General comments (non-PDF):", ApproverComments.length);
    console.log("PDF annotations prepared:", pdfCommentEntries.length);
    console.log("Final comments payload:", allComments);

    let atrType = 1;
    if (selectedAssigneeUsersinfo?.length > 0) {
      atrType = selectedATRType === 2 ? 2 : 3;
    }

    const loadingObj = {
      noteApproverId: currentObj[0]?.noteApproverId,
      noteId: currentObj[0]?.noteId,
      ReferrerEmail: referredEmail,
      approverType: currentObj[0]?.approverType,
      approverEmail: currentObj[0]?.approverEmail,
      approverOrder: currentObj[0]?.approverOrder,
      approverStatus:
        actionBtn === "Approve"
          ? 3
          : actionBtn === "Reject"
            ? 4
            : actionBtn === "Refer"
              ? 6
              : actionBtn === "Return"
                ? 5
                : 1,
      createdDate: new Date().toISOString(),
      createdBy: accounts[0].username,
      modifiedDate: new Date().toISOString(),
      modifiedBy: accounts[0].username,
      strApproverType: currentObj[0]?.strApproverType,
      strApproverStatus:
        actionBtn === "Approve"
          ? "Approved"
          : actionBtn === "Reject"
            ? "Rejected"
            : actionBtn === "Refer"
              ? "Referred"
              : actionBtn === "Return"
                ? "Returned"
                : "",
      noteApproverCommentsDTO: allComments,
      noteSupportingDocumentsDTO:
        actionBtn === "Refer" && totalFileSize > maxFileSize ? [] : fileupdated,
      atrAssigneesDTO: actionBtn === "Approve" ? selectedAssignees : [],
      AtrType: atrType,
    };

    const finalPayload = removeEmptyValues(loadingObj);

    console.log("Final API Payload:", JSON.stringify(finalPayload, null, 2));

    const accessToken = await getAccessToken(
      { ...loginRequest, account },
      instance
    );

    try {
      const response = await fetch(
        `${API_BASE_URL}${API_ENDPOINTS.eNote_NoteApproverStatusChange}`,
        {
          method: "POST",
          headers: {
            ...API_COMMON_HEADERS,
            Authorization: `Bearer ${accessToken}`,
            "Content-Type": "application/json",
          },
          body: JSON.stringify(finalPayload),
        }
      );

      const rep = await response.json();
      SetIsLoading(false);

      if (rep.statusMessage === "Failed") {
        setDialogFailure(true);
        setStatusMessage(rep.statusMessage);
        faildDialogToggle();
      } else {
        if (pdfAnnotatorRef.current?.clearAnnotations) {
          pdfAnnotatorRef.current.clearAnnotations();
        }

        clearCurrentUserDraftAnnotations();
        setPdfAnnotationComments([]);

        let successMessage = "";
        if (actionBtn === "Approve") {
          successMessage = `The request for enote has been ${isApprover() ? "noted" : "approved"
            } successfully.`;
        } else if (actionBtn === "Reject") {
          successMessage = "The request for enote has been rejected successfully.";
        } else if (actionBtn === "Refer") {
          successMessage = "The request for enote has been referred successfully.";
        } else if (actionBtn === "Return") {
          successMessage = "The request for enote has been returned successfully.";
        }

        SetSuccessMsg(successMessage);
        SuccessDialogToggle();
      }
    } catch (err) {
      console.error("API Error:", err);
      SetIsLoading(false);
      faildDialogToggle();
    }
  };
  // ========== PDF ANNOTATION CHANGE END ==========









  const onReferredback = async () => {
    setActionInProgress(true);
    SetIsLoading(true);
    const referComments = [];
    const referredDocs = [];
    let totalFileSize = 0;
    const curntReferredBackUserObj = noteData.noteReferrerDTO?.filter(
      (obj) => obj.referrerEmail === currentActionerEmail
    );
    // Find the referred user with referrerStatus 1 or 3, else fallback to last
    let approverEmail =
      curntReferredBackUserObj[curntReferredBackUserObj.length - 1]?.approverEmail;
    const specialReferrer = curntReferredBackUserObj.find(
      (obj) => obj.referrerStatus === 1 || obj.referrerStatus === 3
    );
    if (specialReferrer) {
      approverEmail = specialReferrer.approverEmail;
    }
    const currentApproverInfo = noteData.noteApproversDTO?.filter(
      (obj) => obj.approverEmail === approverEmail
    );
    // get all Comments Info
    generalcmtforAllCmt?.map((obj) => {
      referComments.push({
        pageNumber: obj.PageNo,
        docReferrence: obj.DocRef,
        referrerComment: obj.Comments,
        NoteReferrerCommentId: obj.noteReferrerCommentID || 0, // Use the actual ID per comment
      });
      return obj;
    });

    const pdfComments =
  pdfAnnotatorRef.current?.getAllFormattedForAPI?.() || [];

pdfComments.forEach((obj) => {
  referComments.push({
    pageNumber: obj.pageNumber || "NA",
    docReferrence: obj.docReferrence || "NA",
    referrerComment: obj.comments || "",
    NoteReferrerCommentId: 0,
    annotationId: obj.annotationId || obj.AnnotationId || null,
  });
});

    //get invaild file name
    const inValidFileNames = Object.keys(supportDocWarning).filter(
      (fileName) => supportDocWarning[fileName]?.isValid === false
    );
    // filter valid files only
    const vaildMultiplefile = supportDocfilesInfo.filter((obj) => {
      if (!inValidFileNames.includes(obj.supportingDocumentFileName)) {
        return obj;
      }
    });
    vaildMultiplefile.map(
      (obj) =>
        (totalFileSize = totalFileSize + obj.supportingDocumentPathLength)
    );
    let fileupdated = vaildMultiplefile?.map((item) => {
      const { supportingDocumentPathLength: _, ...rest } = item;
      return rest;
    });
    // filter valid files only
    supportDocfilesInfo.filter((obj) => {
      if (!inValidFileNames.includes(obj.supportingDocumentFileName)) {
        referredDocs?.push({
          supportingDocumentPath: obj.supportingDocumentPath,
          supportingDocumentFileName: obj.supportingDocumentFileName,
        });
      }
      return obj;
    });

    const updatingObj = removeEmptyValues({
      noteId: noteData.noteId,
      approverType: currentApproverInfo[0]?.approverType,
      approverEmail: approverEmail,
      referrerStatus: 2,
      createdBy: accounts[0].username,
      noteReferrerCommentDTO: referComments.map((comment) => ({
        ...comment,
        noteReffererCommentsId: comment.noteReffererCommentsId,
      })),
      noteSupportingDocumentsDTO: fileupdated,
    });

    const accessToken = await getAccessToken(
      { ...loginRequest, account },
      instance
    );
    await fetch(
      `${API_BASE_URL}${API_ENDPOINTS.eNote_UpdateNoteReferresStatus}`,
      {
        method: "POST",
        headers: {
          ...API_COMMON_HEADERS,
          Authorization: `Bearer ${accessToken}`,
        },
        body: JSON.stringify(updatingObj),
      }
    )
      .then(async (data) => {
        const resp = await data.json();
        if (resp.statusMessage === "Failed") {
          SetIsLoading(false);
          faildDialogToggle();
        } else {
          SetIsLoading(false);
          SetSuccessMsg(
            "The request for enote has been referred back successfully."
          );
          SuccessDialogToggle();
          SetMarkInfoUserEmail([]);
        }
      })
      .catch((err) => {
        SetIsLoading(false);
        faildDialogToggle();
        console.log(err);
      });
  };

  //adding and updating gist Doc attachments details

  const addAttachmentsGistDoc = async () => {
    setActionInProgress(true);
    try {
      SetIsLoading(true);

      const cuurentSecretaryObj = noteData?.noteSecretaryDTO.find(
        (item) => item.secretaryEmail === accounts[0].username
      );
      console.log(cuurentSecretaryObj, "cuurentSecretaryObjcuurentSecretaryObj");
      // Added for base 64 split params ---> Kavya(18-07)
      const base64Parts = wordandPdfInfo.wordInfo.filePath;
      const attchmentDocObj = {
        noteSecretarieId: cuurentSecretaryObj?.noteSecretarieId,
        noteApproverId: cuurentSecretaryObj?.noteApproverId,
        noteId: noteData.noteId,
        secretaryEmail: accounts[0].username,
        approverEmail: noteData?.currentActioner,
        gistWordDocumentFileName: wordandPdfInfo.wordInfo.fileName,
        createdDate: new Date(),
        createdBy: accounts[0].username,
      };

      // Dynamically add the GistWordDocumentPathPart keys -- Kavya (22/07)
      for (let i = 1; i <= 10; i++) {
        attchmentDocObj[`GistWordDocumentPathPart${i}`] = base64Parts[`part${i}`];
      }

      const attachmentDocObjwithoutNull = removeEmptyValues(attchmentDocObj);

      const accessToken = await getAccessToken(
        { ...loginRequest, account },
        instance
      );

      await fetch(`${API_BASE_URL}${API_ENDPOINTS.eNote_InserNoteSecretary}`, {
        method: "POST",
        headers: {
          ...API_COMMON_HEADERS,
          Authorization: `Bearer ${accessToken}`,
        },
        body: JSON.stringify(attachmentDocObjwithoutNull),
      })
        .then(async (data) => {
          const resp = await data.json();

          if (resp.statusMessage === "Failed") {
            SetIsLoading(false);
            faildDialogToggle();
          } else {
            SetIsLoading(false);
            SetSuccessMsg("The request has been updated successfully.");
            SuccessDialogToggle();
            SetMarkInfoUserEmail([]);
          }
        })
        .catch((err) => {
          SetIsLoading(false);
          faildDialogToggle();
          console.log(err);
        });
    } finally {
      setActionInProgress(false);
    }
  };


  // mark info  row intable
  // mark info  row intable
  const addUserInForMarkInfo = () => {
    // SetIsLoading(true);
    const userSlength = markInfoUserEmail?.length;
    if (userSlength <= 9) {
      if (
        selectedMarkUserInfo === "" ||
        selectedMarkUserInfo === undefined ||
        selectedMarkUserInfo === null
      ) {
        setUserNotifyInfo("Please select the user then click on Add User.");
        setUserComboValidationDialog(true);
      } else {
        const ismarkedEmailExist = markInfoUserEmail.some(
          (obj) => obj.markedEmail === selectedMarkUserInfo.userPrincipalName
        );

        if (ismarkedEmailExist) {
          setUserNotifyInfo(
            "The selected user already exist. Kindly choose another user."
          );
          setUserComboValidationDialog(true);
        } else {
          const newMarkInfoUserEmail = [
            ...markInfoUserEmail,
            {
              noteId: noteData.noteId,
              markedEmail: selectedMarkUserInfo.userPrincipalName,
              createdBy: accounts[0].username,
              markedEmailName: selectedMarkUserInfo.displayName,
            },
          ];

          SetMarkInfoUserEmail(newMarkInfoUserEmail);
          setSelectedMarkUserInfo(null);
        }
      }
    } else {
      setUserNotifyInfo("Maximum 10 users allowed.");
      setUserComboValidationDialog(true);
    }
  };

  // remove user for mark info
  const removeUserFromMarkInfo = (id) => {
    SetMarkInfoUserEmail(markInfoUserEmail.filter((obj, ind) => ind !== id));
  };

  // Adding and updating markInfo users details
  const addmarkInfotoDataBase = async () => {
    SetIsLoading(true);
    let markinfouserDetails = markInfoUserEmail?.map((item) => {
      const { markedEmailName: _, ...rest } = item; // Destructure 'markedEmailName' and collect the rest
      return rest; // Return object without 'markedEmailName' property
    });

    const accessToken = await getAccessToken(
      { ...loginRequest, account },
      instance
    );

    await fetch(
      `${API_BASE_URL}${API_ENDPOINTS.eNote_AddNoteMarkedInformation}`,
      {
        method: "POST",
        headers: {
          ...API_COMMON_HEADERS,
          Authorization: `Bearer ${accessToken}`,
        },
        body: JSON.stringify(markinfouserDetails),
      }
    )
      .then(async (data) => {
        const rep = await data.json();
        SetIsLoading(false);
        if (rep.statusMessage === "Failed") {
          faildDialogToggle();
        } else {
          SetSuccessMsg(
            "The mark for information has been updated successfully."
          );
          SuccessDialogToggle();
          SetMarkInfoUserEmail([]);
        }
      })
      .catch((err) => {
        SetIsLoading(false);
        faildDialogToggle();
        console.log(err);
      });
  };

  // const handlecall back function
  const handlecallbackAPICal = async () => {
    SetIsLoading(true);
    let params = {
      noteId: id.id,
      createdBy: accounts[0].username,
    };
    const accessToken = await getAccessToken(
      { ...loginRequest, account },
      instance
    );

    await fetch(`${API_BASE_URL}${API_ENDPOINTS.eNote_CallBackNote}`, {
      method: "POST",
      headers: {
        ...API_COMMON_HEADERS,
        Authorization: `Bearer ${accessToken}`,
      },
      body: JSON.stringify(params),
    })
      .then(async (data) => {
        const res = await data.json();
        SetIsLoading(false);
        if (res.statusMessage === "Failed") {
          SetIsLoading(false);
          faildDialogToggle();
        } else {
          SetIsLoading(false);
          SetSuccessMsg("The request for call back has been successfull.");
          SuccessDialogToggle();
        }
      })
      .catch((err) => {
        faildDialogToggle();
        console.log(err);
      });
  };
  // adding callBack details -this is only availble for Requester
  const handleCallBack = async () => {
    // Added passcode redirect page --> Kavya (23/07) feedback -- 427
    setActionInProgress(true);
    if (verifyPasscode === "true") {
      setPasscodeVerification(true);
      setPasscode("");
      setActionBtn("Call Back");
    } else {
      setValidPasscode(true);
      // Changed the content --> Kavya (25-07)
      setValidmsg(
        "Passcode is not set. Please create passcode to proceed further."
      );
    }
  };

  // handle changeApprover Api cal
  const hnadleChangeApproverAPICall = async () => {
    setActionInProgress(true);
    SetIsLoading(false);
    let params = removeEmptyValues({
      noteId: id.id,
      createdBy: accounts[0].username,
      approverEmail: changeApprovercombovalue.userPrincipalName,
      NoteWordPath: changeApproverhaveSecretary
        ? noteWordDoc?.noteWordDocInfo.filePath
        : null,
      noteWordFileName: changeApproverhaveSecretary
        ? noteWordDoc?.noteWordDocInfo.fileName
        : null,
    });
    // uploadFiles(fileQueue);
    const accessToken = await getAccessToken(
      { ...loginRequest, account },
      instance
    );
    // Make API call
    await fetch(`${API_BASE_URL}${API_ENDPOINTS.eNote_ChangeApprover}`, {
      method: "POST",
      headers: {
        ...API_COMMON_HEADERS,
        Authorization: `Bearer ${accessToken}`,
      },
      body: JSON.stringify(params),
    })
      .then(async (data) => {
        const res = await data.json();
        if (res.statusMessage === "Failed") {
          SetIsLoading(false);
          //Change - Adding SR No - 08/04 - RK
          SetSuccessMsg(`Something went wrong  ${res.statusMessage}`);
          faildDialogToggle();
        } else {
          SetIsLoading(false);
          // Commented for passcode functionality
          setChangeApproverVisible(true);
          setDialogForRefereeAndChangeApprover(false);

        }
      })

      .catch((err) => {
        console.log(err);
        SetIsLoading(false);
        faildDialogToggle();
      });
  };
  //Adding /updating Change Approver details
  const onSubmittingChangeApproverDetails = async () => {
    let isChangeApproverhaveSec = true;

    if (changeApproverhaveSecretary) {
      if (noteWordDoc?.noteWordDocInfo.isValid) {
        isChangeApproverhaveSec = true;
      } else {
        isChangeApproverhaveSec = false;
      }
    }

    if (changeApprovercombovalue && isChangeApproverhaveSec) {
      // Added passcode redirect page --> Kavya (23/07) feedback -- 427
      if (verifyPasscode === "true") {
        setPasscodeVerification(true);
        setPasscode("");
        setActionBtn("ChangeApprover");
      } else {
        setDialogForRefereeAndChangeApprover(false);
        setValidPasscode(true);
        // Changed the content --> Kavya (25-07)
        setValidmsg(
          "Passcode is not set. Please create passcode to proceed further."
        );
      }
    } else {
      setDialogForRefereeAndChangeApprover(false);
      setUserNotifyInfo("Please fill up all the mandatory fields.");
      setUserComboValidationDialog(true);
    }
  };
  // redirect page for passcode --> feedback 427 Kavya (23/07)
  const handleRedirectToPasscode = () => {
    setPasscodeNavigate("View");
    navigate("/enotepasscode"); // Adjust the path to your target route
  };

  // Handle key down (Enter or Space) to show passcode
  const handleKeyDown = (event) => {
    if (event.key === 'Enter' || event.key === ' ') {
      setIsPasscodeVisible(true); // Show passcode
    }
  };

  // Handle key up to hide passcode
  const handleKeyUp = (event) => {
    if (event.key === 'Enter' || event.key === ' ') {
      setIsPasscodeVisible(false); // Hide passcode
    }
  };

  const handleATRTypeChange = async (e) => {
    const selectedType = e.value;
    setSelectedATRType(selectedType); // Set the selected ATR type
    setSelectedAssigneeUsersinfo([]); // Clear the selected assignee users
    setSelectedAssignee(null); // Clear the selected assignee
    if (selectedType === 2) { // Check for internal selection

      try {
        SetIsLoading(true);

        const accessToken = await getAccessToken(
          { ...loginRequest, account },
          instance
        );
        const params = {
          noteId: id.id,
          LoginMailID: accounts[0].username,
        };

        const response = await fetch(`${API_BASE_URL}${API_ENDPOINTS.atrInternalDropdeown}`, {
          method: "POST",
          headers: {
            ...API_COMMON_HEADERS,
            Authorization: `Bearer ${accessToken}`,
          },
          body: JSON.stringify(params),
        });

        const res = await response.json();
        setInternalUsers(res); // Set internal users
        if (res.statusMessage === "Failed") {
          SetSuccessMsg(`Something went wrong: ${res.statusMessage}`);
          faildDialogToggle();
        }
      } catch (error) {
        console.error("API request failed", error);
        SetSuccessMsg("An error occurred. Please try again.");
      } finally {
        SetIsLoading(false);
      }
    } else if (selectedType === 3) {
      // If external is selected, use the existing external data
      setAssigneeUsers(assigneerUserDetails);
    }
  };

  return (
    <div>
      {showProgressBar && <LoadingOverlay value={progressValue} />}
      <Navbar isViewPage />
      <Sidebar />
      {/* Added top navigation for menu --> 23-09 */}
      <MenuNavContainer isViewPage />
      <>
        {isVieweEligible === "false" ? (
          <>
            {" "}
            <Unauthorized />{" "}
            <div className="pgFooterContainer">
              <Footer />
            </div>{" "}
          </>
        ) : (
          <>
            {" "}
            <div className="FormMaincontainer viewformContainer">
              <div className="container largeScreen" >
                <fieldset className={"k-form-fieldset"}>
                  <div>
                    <div className="viewHeader">
                      <div className="SectionHeads-viewForm">
                        <span className="cstformHdrLeftContainer mobiletitle">
                          {`Pending with: ${noteData?.currentActionerName || ""}`}
                        </span>
                        <span className="cstformHdrMiddleContainer mobiletitle">
                          eNote - {noteData?.noteNumber || ""}
                        </span>
                        <span className="cstformHdrRightContainer mobiletitle">
                          {`Status : ${noteData ? noteData?.strNoteStatus : ""}`}
                        </span>
                      </div>
                    </div>
                    <div className="viewFormhomepageContainer">
                      {/* Change 05/04 adding dynamic class for hide and show */}
                      <div
                        ref={expansionPanelRef}
                        className={
                          isPDFFullWidth ? "homesectionPdf-1" : "viewFormSection-1"
                        }
                      >
                        {expendJson.map((item) =>
                          item.isVisible ? (
                            <ExpansionPanel
                              title={
                                <div className="_ExpansionPanel">
                                  <div />
                                  {item.id}
                                </div>
                              }
                              expanded={expanded.includes(item.id)}
                              tabIndex={0}
                              onAction={() => handlePanelToggle(item.id)}
                            >
                              <Reveal>
                                {noteData !== null && (
                                  <>
                                    {expanded.includes(item.id) && (
                                      <ExpansionPanelContent>
                                        {item.id === "General Section" &&
                                          item.isVisible && (
                                            <>
                                              {/* TAT Button for Admins */}
                                              {/* {(deptInfo?.isSuperAdmin === true || deptInfo?.isDeptAdmin === true) && ( */}
                                              <Button className="formBtnColor tatButton" onClick={tatDialogToggle}>
                                                <span>TAT</span>
                                              </Button>
                                              {/* )} */}

                                              {tatDialog && (
                                                <Dialog
                                                  title={<div className="k-dialog-titlebar"><span className="k-dialog-title">TAT Report</span></div>}
                                                  onClose={tatDialogToggle}
                                                  className="dialog"
                                                >
                                                  <div style={{ marginTop: "10px", maxHeight: "500px", overflowY: "auto" }}>
                                                    <Grid data={buildTatData(noteData)} style={{ maxHeight: "450px" }}>
                                                      {/* <Column field="office" title="Office" width="180px" /> */}
                                                      <Column field="steps" title="Steps" width="350px" />
                                                      <Column field="submittedOn" title="Received On" width="130px" />
                                                      <Column field="approvedOn" title="Action Taken On" width="150px" />
                                                      <Column field="individualTat" title="Individual TAT (Hours)" width="150px" />
                                                      <Column field="officeTat" title="office Tat" width="130px" />
                                                    </Grid>
                                                  </div>

                                                  <DialogActionsBar>
                                                    <div className="k-dialog-actions k-actions k-actions-center k-actions-horizontal">
                                                      <Button className="k-button k-button-md k-rounded-md k-button-solid k-button-solid-base" onClick={tatDialogToggle}>
                                                        <span className="k-icon k-font-icon k-i-close-circle cursor allIconsforPrimary-btn"></span>
                                                        <span>Cancel</span>
                                                      </Button>
                                                    </div>
                                                  </DialogActionsBar>
                                                </Dialog>
                                              )}
                                              {" "}

                                              {noteData && (
                                                <table className="ib-forms-custom-table tableStyle">
                                                  <tbody>
                                                    {/* TAT Button for Admins */}
                                                    {/* {(deptInfo?.isSuperAdmin === true || deptInfo?.isDeptAdmin === true) && ( */}

                                                    {/* )} */}
                                                    <tr>
                                                      <th className="approvalform-tableCol-width-5 generalPadding">
                                                        Note Number
                                                      </th>
                                                      <td className="approvalform-tableCol-width-11 generalPadding">
                                                        {noteData.noteNumber}
                                                      </td>
                                                    </tr>
                                                    <tr>
                                                      <th className="approvalform-tableCol-width-5 generalPadding">
                                                        Requester{" "}
                                                      </th>
                                                      <td className="approvalform-tableCol-width-11 generalPadding">
                                                        {noteData.createdByName}
                                                      </td>
                                                    </tr>
                                                    <tr>
                                                      <th className="approvalform-tableCol-width-5 generalPadding">
                                                        Request Date{" "}
                                                      </th>
                                                      <td className="approvalform-tableCol-width-11 generalPadding">
                                                        {getdateconversion(
                                                          noteData.createdDate
                                                        )}
                                                      </td>
                                                    </tr>
                                                    <tr>
                                                      <th className="approvalform-tableCol-width-5 generalPadding">
                                                        Status{" "}
                                                      </th>
                                                      <td className="approvalform-tableCol-width-11 generalPadding">
                                                        {noteData.strNoteStatus}
                                                      </td>
                                                    </tr>
                                                    <tr>
                                                      <th className="approvalform-tableCol-width-5 generalPadding">
                                                        Note to{" "}
                                                      </th>
                                                      <td className="approvalform-tableCol-width-11 generalPadding">
                                                        {noteData.strNoteTo}
                                                      </td>
                                                    </tr>
                                                    <tr>
                                                      <th className="approvalform-tableCol-width-5 generalPadding">
                                                        Department{" "}
                                                      </th>
                                                      <td className="approvalform-tableCol-width-11 generalPadding">
                                                        {noteData.departmentName}
                                                      </td>
                                                    </tr>
                                                    <tr>
                                                      <th className="approvalform-tableCol-width-5 generalPadding">
                                                        Subject
                                                      </th>
                                                      <td className="approvalform-tableCol-width-11 generalPadding">
                                                        {noteData.subject}
                                                      </td>
                                                    </tr>
                                                    <tr>
                                                      <th className="approvalform-tableCol-width-5 generalPadding">
                                                        Nature of the Note{" "}
                                                      </th>
                                                      <td className="approvalform-tableCol-width-11 generalPadding">
                                                        {noteData.strNatureofNote}
                                                      </td>
                                                    </tr>
                                                    {(noteData.strNatureofNote === "Sanction" ||
                                                      noteData.strNatureofNote === "Approval") && (
                                                        <tr>
                                                          <th className="approvalform-tableCol-width-5 generalPadding">
                                                            Nature of Approval /
                                                            Sanction{" "}
                                                          </th>
                                                          <td className="approvalform-tableCol-width-11 generalPadding">
                                                            {noteData.strNatureOfApprovalOrSanction ===
                                                              "0"
                                                              ? "NA"
                                                              : noteData.strNatureOfApprovalOrSanction}
                                                          </td>
                                                        </tr>
                                                      )}

                                                    <tr>
                                                      <th className="approvalform-tableCol-width-5 generalPadding">
                                                        Note Type{" "}
                                                      </th>
                                                      <td className="approvalform-tableCol-width-11 generalPadding">
                                                        {noteData.strNoteType}
                                                      </td>
                                                    </tr>
                                                    {noteData.strNoteType ===
                                                      "Financial" && (
                                                        <tr>
                                                          <th className="approvalform-tableCol-width-5 generalPadding">
                                                            Amount
                                                          </th>
                                                          <td className="approvalform-tableCol-width-11 generalPadding">
                                                            {noteData.strAmount}
                                                          </td>
                                                        </tr>
                                                      )}
                                                    {/* <tr>
                                                      <th className="approvalform-tableCol-width-5 generalPadding">
                                                        Search Keyword{" "}
                                                      </th>
                                                      <td className="approvalform-tableCol-width-11 generalPadding">
                                                        {noteData.searchKeyword}
                                                      </td>
                                                    </tr>
                                                    <tr>
                                                      <th className="approvalform-tableCol-width-5 generalPadding">
                                                        Purpose
                                                      </th>
                                                      <td className="approvalform-tableCol-width-11 generalPadding">
                                                        {noteData.purpose}
                                                      </td>
                                                    </tr> */}
                                                    {noteData.isConfidential === true && (
                                                      <tr>
                                                        <th className="approvalform-tableCol-width-5 generalPadding">
                                                          Confidential Note:
                                                        </th>
                                                        <td className="approvalform-tableCol-width-11 generalPadding">
                                                          {noteData.isConfidential === true ? "Yes" : "No"}
                                                        </td>
                                                      </tr>
                                                    )}
                                                  </tbody>
                                                </table>
                                              )}
                                            </>
                                          )}
                                        {item.id === "Reviewers Section" &&
                                          item.isVisible && (
                                            <>
                                              {noteData && (
                                                <div className="table-responsive">
                                                  <table className="ib-forms-custom-table tableStyle">
                                                    <tbody>
                                                      <tr>
                                                        {/* Bug fix - 294 - 27/03 */}
                                                        <th className="approvalform-tableCol-width-2">Reviewer Name{" "}</th>
                                                        <th className="approvalform-tableCol-width-1">Employee Id</th>
                                                        <th className="approvalform-tableCol-width-2">Designation</th>
                                                        <th className="approvalform-tableCol-width-2"> Status{" "}</th>
                                                        <th className="approvalform-tableCol-width-3">Action Date{" "}</th>
                                                      </tr>
                                                      {noteData.noteApproversDTO?.map(
                                                        (obj) =>
                                                          obj.approverType === 1 && obj.isActive && (
                                                            <tr
                                                              key={obj.approverEmail}
                                                              className={`${obj.approverStatus === 2 ? "selected-row" : ""}`}
                                                            >
                                                              <td className="approvalform-tableCol-width-2">
                                                                {obj.approverEmailName}
                                                              </td>
                                                              <td className="approvalform-tableCol-width-1">
                                                                {obj.srNo}
                                                              </td>
                                                              <td className="approvalform-tableCol-width-2">
                                                                {obj.designation}
                                                              </td>
                                                              <td className="approvalform-tableCol-width-2">
                                                                <span
                                                                  className={`k-icon k-font-icon ${renderApproverIcon(
                                                                    obj.approverStatus
                                                                  )}  allIconsforPrimary-btn`}
                                                                ></span>
                                                                {obj.strApproverStatus}
                                                              </td>
                                                              <td className="approvalform-tableCol-width-3">
                                                                {obj.approverStatus === 1 || obj.approverStatus === 2 ? ""
                                                                  : getdateconversion(obj.modifiedDate)}
                                                              </td>
                                                            </tr>
                                                          )
                                                      )}
                                                    </tbody>
                                                  </table>
                                                </div>
                                              )}
                                            </>
                                          )}
                                        {item.id === "Approvers Section" &&
                                          item.isVisible && (
                                            <div className="table-responsive">
                                              {" "}
                                              <table className="ib-forms-custom-table tableStyle">
                                                <tbody></tbody>{" "}
                                                <tr>
                                                  <th className="approvalform-tableCol-width-2">
                                                    Approver Name{" "}
                                                  </th>
                                                  <th className="approvalform-tableCol-width-1">
                                                    Employee Id
                                                  </th>
                                                  <th className="approvalform-tableCol-width-2">
                                                    Designation
                                                  </th>
                                                  <th className="approvalform-tableCol-width-2">
                                                    Status{" "}
                                                  </th>
                                                  <th className="approvalform-tableCol-width-3">
                                                    Action Date{" "}
                                                  </th>
                                                </tr>
                                                {noteData.noteApproversDTO?.map(
                                                  (obj) =>
                                                    obj.approverType === 2 && obj.isActive && (
                                                      <tr
                                                        key={obj.approverEmail}
                                                        className={`${obj.approverStatus === 2
                                                          ? "selected-row"
                                                          : ""
                                                          }`}
                                                      >

                                                        <td className="approvalform-tableCol-width-2">
                                                          {obj.approverEmailName}
                                                        </td>
                                                        <td className="approvalform-tableCol-width-1">
                                                          {obj.srNo}
                                                        </td>
                                                        <td className="approvalform-tableCol-width-2">
                                                          {obj.designation}
                                                        </td>
                                                        <td className="approvalform-tableCol-width-2">
                                                          <span
                                                            className={`k-icon k-font-icon ${renderApproverIcon(
                                                              obj.approverStatus
                                                            )}  allIconsforPrimary-btn`}
                                                          ></span>
                                                          {obj.strApproverStatus}
                                                        </td>
                                                        <td className="approvalform-tableCol-width-3">
                                                          {obj.approverStatus === 1 ||
                                                            obj.approverStatus === 2
                                                            ? ""
                                                            : getdateconversion(
                                                              obj.modifiedDate
                                                            )}
                                                        </td>
                                                      </tr>
                                                    )
                                                )}
                                              </table>
                                            </div>
                                          )}
                                        {item.id === "General Comments" &&
                                          item.isVisible && (
                                            <div>
                                              <div className="table-responsive">
                                                <table className="ib-forms-custom-table tableStyle mob_table">
                                                  <thead>
                                                    <tr>
                                                      <th className="approvalform-tableCol-width-1">
                                                        Page#
                                                      </th>
                                                      <th className="approvalform-tableCol-width-1">
                                                        Doc Reference
                                                      </th>
                                                      <th className="approvalform-tableCol-width-6 mobComments">
                                                        Comments
                                                      </th>
                                                      <th className="approvalform-tableCol-width-2">
                                                        Action
                                                      </th>
                                                    </tr>
                                                  </thead>
                                                  <tbody>
                                                    <tr className="non_mobileResponsive">
                                                      <td className="approvalform-tableCol-width-1">
                                                        <div
                                                          onKeyDown={(e) =>
                                                            e.stopPropagation()
                                                          }
                                                        >
                                                          <input
                                                            className="general-comments-pageNo"
                                                            type="text"
                                                            name="PageNo"
                                                            value={
                                                              generalcmtInfoobj.PageNo
                                                            }
                                                            onChange={
                                                              handelGenralComments
                                                            }
                                                          />
                                                        </div>
                                                      </td>
                                                      <td className="approvalform-tableCol-width-1">
                                                        <div
                                                          onKeyDown={(e) =>
                                                            e.stopPropagation()
                                                          }
                                                        >
                                                          <input
                                                            className="general-comments-DocRef"
                                                            name="DocRef"
                                                            value={
                                                              generalcmtInfoobj.DocRef
                                                            }
                                                            onChange={
                                                              handelGenralComments
                                                            }
                                                          />
                                                        </div>
                                                      </td>
                                                      <td className="approvalform-tableCol-width-6">
                                                        <div
                                                          onKeyDown={(e) =>
                                                            e.stopPropagation()
                                                          }
                                                        >
                                                          <textarea
                                                            name="Comments"
                                                            maxLength={max}
                                                            className="general-comments-Comments mobComments"
                                                            row={2}
                                                            value={
                                                              generalcmtInfoobj.Comments
                                                            }
                                                            onChange={
                                                              handelGenralComments
                                                            }
                                                            onClick={
                                                              handleCommentsMobile
                                                            }
                                                            disabled={
                                                              mobileCommentsDialog
                                                            } // This should be true when dialog is open
                                                          ></textarea>
                                                          <Hint direction={"end"}>
                                                            {generalcmtInfoobj.Comments.length} / {max}
                                                          </Hint>
                                                        </div>
                                                      </td>
                                                      <td className="approvalform-tableCol-width-2">
                                                        <Button
                                                          className="formBtnColor non_mobileResponsive"
                                                          onClick={handelAddCmts}
                                                        >
                                                          <span className="k-icon k-font-icon k-i-plus cursor allIconsforPrimary-btn"></span>
                                                          Add
                                                        </Button>
                                                      </td>
                                                    </tr>
                                                    {generalcmtforAllCmt?.map(
                                                      (obj, index) => (
                                                        <tr key={index}>
                                                          <td className="approvalform-tableCol-width-1">
                                                            {obj.isEdit ? (
                                                              <div
                                                                onKeyDown={(e) =>
                                                                  e.stopPropagation()
                                                                }
                                                              >
                                                                <input
                                                                  type="text"
                                                                  className="general-comments-pageNo"
                                                                  value={obj.PageNo}
                                                                  name="PageNo"
                                                                  id={index}
                                                                  onChange={
                                                                    handelGenralCommentsEdit
                                                                  }
                                                                  onClick={() =>
                                                                    handleCommentsEditMobile(
                                                                      obj.Comments,
                                                                      obj.PageNo,
                                                                      obj.DocRef,
                                                                      index
                                                                    )
                                                                  }
                                                                />

                                                              </div>
                                                            ) : (
                                                              obj.PageNo
                                                            )}
                                                          </td>
                                                          <td className="approvalform-tableCol-width-1">
                                                            {obj.isEdit ? (
                                                              <div
                                                                onKeyDown={(e) =>
                                                                  e.stopPropagation()
                                                                }
                                                              >
                                                                <input
                                                                  type="text"
                                                                  className="general-comments-DocRef"
                                                                  value={obj.DocRef}
                                                                  name="DocRef"
                                                                  id={index}
                                                                  onChange={
                                                                    handelGenralCommentsEdit
                                                                  }
                                                                  onClick={() =>
                                                                    handleCommentsEditMobile(
                                                                      obj.Comments,
                                                                      obj.PageNo,
                                                                      obj.DocRef,
                                                                      index
                                                                    )
                                                                  }
                                                                />
                                                              </div>
                                                            ) : (
                                                              obj.DocRef
                                                            )}
                                                          </td>
                                                          <td className="approvalform-tableCol-width-6">
                                                            {obj.isEdit ? (
                                                              <div
                                                                onKeyDown={(e) =>
                                                                  e.stopPropagation()
                                                                }
                                                              >
                                                                <textarea
                                                                  value={obj.Comments}
                                                                  className="general-comments-Comments"
                                                                  name="Comments"
                                                                  id={index}
                                                                  onChange={
                                                                    handelGenralCommentsEdit
                                                                  }
                                                                  onClick={() =>
                                                                    handleCommentsEditMobile(
                                                                      obj.Comments,
                                                                      obj.PageNo,
                                                                      obj.DocRef,
                                                                      index
                                                                    )
                                                                  }
                                                                ></textarea>
                                                                <Hint direction={"end"}>
                                                                  {obj.Comments.length} / {max}
                                                                </Hint>
                                                              </div>
                                                            ) : (
                                                              <div style={{ whiteSpace: "pre-wrap" }}>{obj.Comments}</div>
                                                            )}
                                                          </td>

                                                          {/* non_mobileResponsive */}
                                                          <td className="approvalform-tableCol-width-2">
                                                            {obj.isEdit ? (
                                                              <>
                                                                <Button
                                                                  className="formBtnColor non_mobileResponsive"
                                                                  onClick={() =>
                                                                    handleGeneraleditedcommentsSave(
                                                                      index
                                                                    )
                                                                  }
                                                                >
                                                                  <span className="k-icon k-font-icon k-i-save cursor allIconsforPrimary-btn"></span>
                                                                  Save
                                                                </Button>
                                                                <Button
                                                                  className="formBtnColor _mobileResponsive"
                                                                  onClick={() =>
                                                                    handleGeneraleditedcommentsSave(
                                                                      index
                                                                    )
                                                                  }
                                                                >
                                                                  <span className="k-icon k-font-icon k-i-save cursor allIconsforPrimary-btn"></span>
                                                                </Button>
                                                              </>
                                                            ) : (
                                                              <>
                                                                {/* non_mobileResponsive */}
                                                                <Button
                                                                  className="non_mobileResponsive"
                                                                  onClick={() =>
                                                                    handelDeleteCmt(
                                                                      index
                                                                    )
                                                                  }
                                                                >
                                                                  <span className="k-icon k-font-icon k-i-delete k-i-trash cursor allIconsforPrimary-btn"></span>
                                                                  Delete
                                                                </Button>
                                                                <Button
                                                                  className="_mobileResponsive"
                                                                  onClick={() =>
                                                                    handelDeleteCmt(
                                                                      index
                                                                    )
                                                                  }
                                                                >
                                                                  <span className="k-icon k-font-icon k-i-delete k-i-trash cursor allIconsforPrimary-btn"></span>
                                                                </Button>
                                                                <br />
                                                                <Button
                                                                  className="non_mobileResponsive"
                                                                  onClick={() =>
                                                                    handelingEdit(
                                                                      index
                                                                    )
                                                                  }
                                                                >
                                                                  <span className="k-icon k-font-icon k-i-edit cursor allIconsforPrimary-btn"></span>
                                                                  Edit
                                                                </Button>
                                                                <Button
                                                                  className="_mobileResponsive"
                                                                  onClick={() =>
                                                                    handelingEdit(
                                                                      index
                                                                    )
                                                                  }
                                                                >
                                                                  <span className="k-icon k-font-icon k-i-edit cursor allIconsforPrimary-btn"></span>
                                                                </Button>
                                                              </>
                                                            )}
                                                          </td>
                                                        </tr>
                                                      )
                                                    )}
                                                  </tbody>
                                                </table>
                                                {errorMessageEdit && (
                                                  <div className="inCorrectFileError">
                                                    {errorMessageEdit}
                                                  </div>
                                                )}
                                              </div>
                                              <div className="mobileCommentButton">
                                                <Button
                                                  className="formBtnColor _mobileResponsive addMobile"
                                                  onClick={handleOpenCommentDialog}
                                                >
                                                  <span className="k-icon k-font-icon k-i-plus cursor allIconsforPrimary-btn"></span>
                                                  Add
                                                </Button>
                                              </div>
                                            </div>
                                          )}
                                        {item.id === "ATR Assignees" && item.isVisible && (
                                          <>
                                            <div className="atrTypeStyle">
                                              {eNumObj.ATRType
                                                .filter((type) => type.dValue.trim().toLowerCase() !== "default")
                                                .map((type) => (
                                                  <div key={type.id}>
                                                    <RadioButton
                                                      name="ATRTypeGroup"
                                                      value={type.id}
                                                      // Check if the external option should be selected on load
                                                      checked={selectedATRType ? selectedATRType === type.id : type.dValue.trim().toLowerCase() === "external"}
                                                      label={type.DisplayName}
                                                      onChange={handleATRTypeChange}  // Handles the selection change
                                                      className="atrTypeRadioButton"
                                                    />
                                                  </div>
                                                ))}
                                            </div>

                                            <div>
                                              <div className="enote-Assignees-container">
                                                <div onKeyDown={(e) => e.stopPropagation()}>
                                                  {/* Render ComboBox based on selected ATR Type */}
                                                  {selectedATRType === 2 ? (

                                                    <ComboBox
                                                      disabled={
                                                        !noteData?.noteApproversDTO?.some(
                                                          (obj) =>
                                                            obj.approverEmail === accounts[0].username &&
                                                            obj.approverEmail === noteData?.currentActioner &&
                                                            obj.approverType === 2 &&
                                                            obj.approverStatus === 2
                                                        )
                                                      }
                                                      data={
                                                        internalUsers.length === 0
                                                          ? ["NA"]
                                                          : [...new Set(internalUsers.map((obj) => obj.approverEmailName))]
                                                      }
                                                      value={selectedAssignee}
                                                      filterable={true}
                                                      onFilterChange={onFilterChangeForAtrAssigness}
                                                      onChange={handleComboChangeAssignees}
                                                      placeholder="Select Assignee.."
                                                      className="_atrAssineesCombobox"
                                                    />
                                                  ) : (
                                                    <ComboBox
                                                      disabled={
                                                        !noteData?.noteApproversDTO?.some(
                                                          (obj) =>
                                                            obj.approverEmail === accounts[0].username &&
                                                            obj.approverEmail === noteData?.currentActioner &&
                                                            obj.approverType === 2 &&
                                                            obj.approverStatus === 2
                                                        )
                                                      }
                                                      data={[
                                                        ...new Set(
                                                          assigneeUsers.map((obj) => obj.atrAssignerEmailName)
                                                        ),
                                                      ]}
                                                      value={selectedAssignee}
                                                      filterable={true}
                                                      onFilterChange={onFilterChangeForAtrAssigness}
                                                      onChange={handleComboChangeAssignees}
                                                      placeholder="Select Assignee.."
                                                      className="_atrAssineesCombobox"
                                                    />
                                                  )}
                                                </div>

                                                {/* Add buttons */}
                                                <Button
                                                  className="formBtnColor non_mobileResponsive"
                                                  onClick={handleAddRowAssignees}
                                                  disabled={
                                                    !noteData?.noteApproversDTO?.some(
                                                      (obj) =>
                                                        obj.approverEmail === accounts[0].username &&
                                                        obj.approverEmail === noteData?.currentActioner &&
                                                        obj.approverStatus === 2
                                                    )
                                                  }
                                                >
                                                  <span className="k-icon k-font-icon k-i-plus cursor allIconsforPrimary-btn"></span>
                                                  Add
                                                </Button>
                                                <Button
                                                  className="formBtnColor _mobileResponsive"
                                                  onClick={handleAddRowAssignees}
                                                  disabled={
                                                    !noteData?.noteApproversDTO?.some(
                                                      (obj) =>
                                                        obj.approverEmail === accounts[0].username &&
                                                        obj.approverEmail === noteData?.currentActioner &&
                                                        obj.approverStatus === 2
                                                    )
                                                  }
                                                >
                                                  <span className="k-icon k-font-icon k-i-plus cursor allIconsforPrimary-btn"></span>

                                                </Button>
                                              </div>

                                              {/* Table displaying assignees */}
                                              <div className="table-responsive">
                                                <table className="ib-forms-custom-table tableStyle">
                                                  <thead>
                                                    <tr>
                                                      <th className="approvalform-tableCol-width-3">Comments</th>
                                                      <th className="approvalform-tableCol-width-3">Assigned To</th>
                                                      <th className="approvalform-tableCol-width-2">Status</th>
                                                      {noteData?.noteApproversDTO?.some(
                                                        (obj) =>
                                                          obj.approverEmail === accounts[0].username &&
                                                          obj.approverEmail === noteData?.currentActioner &&
                                                          obj.approverStatus === 2
                                                      ) && (
                                                          <th className="approvalform-tableCol-width-2">Action</th>
                                                        )}
                                                    </tr>
                                                  </thead>
                                                  <tbody>
                                                    {noteData?.atrAssigneesDTO?.map(
                                                      (obj, index) =>
                                                        obj.approverEmail === accounts[0].username && (
                                                          <tr key={obj.atrAssignerEmail}>
                                                            <td>{obj.noteApproverComments}</td>
                                                            <td>{obj.atrAssignerEmailName}</td>
                                                            <td>{obj.strATRStatus}</td>
                                                            {noteData?.noteApproversDTO?.some(
                                                              (obj) =>
                                                                obj.approverEmail === accounts[0].username &&
                                                                obj.approverEmail === noteData?.currentActioner &&
                                                                obj.approverStatus === 2
                                                            ) && <td>{/* Action buttons */}</td>}
                                                          </tr>
                                                        )
                                                    )}
                                                    {selectedAssigneeUsersinfo?.map((obj, index) => (
                                                      <tr key={obj.atrAssignerEmail}>
                                                        <td className="approvalform-tableCol-width-3">
                                                          {generalcmtforAllCmt
                                                            ?.map(
                                                              (obj) => `${obj.PageNo} ${obj.DocRef} ${obj.Comments}`
                                                            )
                                                            .join(", ")}
                                                        </td>
                                                        <td className="approvalform-tableCol-width-3">
                                                          {obj?.atrAssignerEmailName}
                                                        </td>
                                                        <td className="approvalform-tableCol-width-2">Submitted</td>
                                                        <td className="approvalform-tableCol-width-2">
                                                          <Button
                                                            className="non_mobileResponsive"
                                                            onClick={() => onDeleteAssigneeUser(index)}
                                                          >
                                                            <span className="k-icon k-font-icon k-i-delete k-i-trash cursor allIconsforPrimary-btn"></span>
                                                            Delete
                                                          </Button>
                                                          <Button
                                                            className="_mobileResponsive"
                                                            onClick={() => onDeleteAssigneeUser(index)}
                                                          >
                                                            <span className="k-icon k-font-icon k-i-delete k-i-trash cursor allIconsforPrimary-btn"></span>
                                                          </Button>
                                                        </td>
                                                      </tr>
                                                    ))}
                                                  </tbody>
                                                </table>
                                              </div>
                                            </div>
                                          </>
                                        )}

                                        {/* {item.id === "Comments Log" &&
                                          item.isVisible && (
                                            <div className="table-responsive">
                                              {" "}
                                              <table className="ib-forms-custom-table tableStyle mob_table">
                                                <tbody>
                                                  {" "}
                                                  <tr>
                                                    <th className="approvalform-tableCol-width-1">
                                                      Page#
                                                    </th>
                                                    <th className="approvalform-tableCol-width-1">
                                                      Doc Reference
                                                    </th>
                                                    <th className="approvalform-tableCol-width-5 mobComments">
                                                      Comments
                                                    </th>
                                                    <th className="approvalform-tableCol-width-3">
                                                      Comment By
                                                    </th>
                                                    <th className="approvalform-tableCol-width-3">
                                                      Comment Date
                                                    </th>
                                                  </tr>
                                                  {noteData.noteApproverCommentsDTO.map(
                                                    (comment) => (
                                                      <tr
                                                        key={
                                                          comment.noteApproverCommentID
                                                        }
                                                      >
                                                        <td className="approvalform-tableCol-width-1">
                                                          {comment.pageNumber || "NA"}
                                                        </td>
                                                        <td className="approvalform-tableCol-width-1">
                                                          {comment.docReferrence ||
                                                            "NA"}
                                                        </td>
                                                        <td className="approvalform-tableCol-width-5">
                                                          {comment.comments}
                                                        </td>
                                                        <td className="approvalform-tableCol-width-3">
                                                          {comment.approverEmailName}
                                                        </td>
                                                        <td className="approvalform-tableCol-width-3">
                                                          {getdateconversion(
                                                            comment.modifiedDate
                                                          )}
                                                        </td>
                                                      </tr>
                                                    )
                                                  )}
                                                  {generalcmtforAllCmt?.slice().reverse().map(
                                                    (obj, index) =>
                                                      !obj.isEdit && (
                                                        <tr key={index}>
                                                          <td className="approvalform-tableCol-width-1">
                                                            {obj.PageNo || "NA"}
                                                          </td>
                                                          <td className="approvalform-tableCol-width-1">
                                                            {obj.DocRef || "NA"}
                                                          </td>
                                                          <td className="approvalform-tableCol-width-5">
                                                            {obj.Comments}
                                                          </td>
                                                          <td className="approvalform-tableCol-width-3">
                                                            {accounts[0].name}
                                                          </td>
                                                        </tr>
                                                      )
                                                  )}
                                                </tbody>{" "}
                                              </table>
                                            </div>
                                          )} */}
                                        {item.id === "Attach Supporting Documents" &&
                                          item.isVisible && (
                                            <>
                                              <div className="SectionRow row">
                                                <div className="col">
                                                  <span className="k-form-label">
                                                    <b>Attach Supporting Documents</b>
                                                  </span>
                                                  <div className="Attachemntfileinfo-ind">
                                                    <input
                                                      type="file"
                                                      id="multiDoc"
                                                      multiple
                                                      key={fileKey} // Unique key to force re-render
                                                      // onChange={(event) => multipleDocUpload(event, 'SupportingDocument')} // Supporting Document as file category
                                                      onChange={multipleDocUpload}
                                                      // onChange={(event) => multipleDocUpload(event.target.files[0], 'SupportingDocument')}
                                                      className="_Attachemntfileinfo_multiple"
                                                    />
                                                    <div>
                                                      <span className="inCorrectFileError">
                                                        {supportingDocError}
                                                      </span>
                                                    </div>
                                                  </div>
                                                  {supportDocfilesInfo?.map(
                                                    (obj, ind) => (
                                                      <div
                                                        className="Attachemntfileinfo-ind"
                                                        key={ind}
                                                      >
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
                                                              <div className="attachemnrt-warningifo-fileinfo attachment_warn">
                                                                {obj.supportingDocumentFileName}
                                                              </div>
                                                              <span className="inCorrectFileError">
                                                                {supportDocWarning[obj.supportingDocumentFileName]?.warningMsg}
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
                                                    )
                                                  )}
                                                  <div className="_fileFormatHintMsg">
                                                    {/* Allowed Formats (images,pdf,eml,doc,docx,xlsx only) Upto 5MB max. */}
                                                    Allowed Formats (pdf,doc,docx,xlsx only) Upto 25MB max.
                                                  </div>
                                                </div>
                                              </div>
                                            </>
                                          )}
                                        {item.id === "Gist Document" &&
                                          item.isVisible && (
                                            <>
                                              <div className="SectionRow row">
                                                <div className="col">
                                                  <span className="k-form-label">
                                                    <b>Gist Document</b>
                                                  </span>
                                                  {
                                                    (noteData?.noteSecretaryDTO?.some(
                                                      (obj) =>
                                                        obj.secretaryEmail === accounts[0].username &&
                                                        obj.approverEmail === noteData?.currentActioner
                                                    ) || (isGistSecretary && noteData?.currentActionerDesignation === MDNameTitle)) && (
                                                      <div className="Attachemntfileinfo-ind">
                                                        <input
                                                          type="file"
                                                          id="WordDocfile"
                                                          key={fileKey} // Unique key to force re-render
                                                          // onChange={(event) => attchemtsForGistDoc(event, 'Supporting Document')}
                                                          // onChange={(event) => attchemtsForGistDoc(event.target.files[0], 'Supporting Document')}
                                                          onChange={attchemtsForGistDoc}
                                                          className="_ConvertPDFToBase64"
                                                        />
                                                      </div>
                                                    )}
                                                  {wordandPdfInfo?.wordInfo
                                                    .fileName === "" ? null : (
                                                    <div className="Attachemntfileinfo-ind">
                                                      <span className="attachemntIconInfoConationer">
                                                        <span className="AttchemntIconWraper">
                                                          <SvgIcon
                                                            icon={getFileIcon(wordandPdfInfo?.wordInfo.fileName)}
                                                            size="xxlarge"
                                                          />
                                                          <span className="attachemnrt-warningifoConatiner">
                                                            <div className="attachemnrt-warningifo-fileinfo">
                                                              {wordandPdfInfo
                                                                ?.wordInfo
                                                                .isDownloadble ? (
                                                                <a
                                                                  href={wordandPdfInfo?.wordInfo.base64}
                                                                  download={wordandPdfInfo.wordInfo.fileName}
                                                                >
                                                                  {wordandPdfInfo?.wordInfo.fileName}
                                                                </a>
                                                              ) : (
                                                                wordandPdfInfo?.wordInfo.fileName
                                                              )}
                                                            </div>
                                                            <span className="inCorrectFileError">{fileWordWarning}</span>
                                                          </span>
                                                        </span>

                                                        {(noteData?.noteSecretaryDTO?.some(
                                                          (obj) =>
                                                            obj.secretaryEmail ===
                                                            accounts[0].username &&
                                                            obj.approverEmail ===
                                                            noteData?.currentActioner) ||
                                                          (isGistSecretary && noteData?.currentActionerDesignation === MDNameTitle)
                                                        ) && (
                                                            <span
                                                              className="AttchemntIconWraperCancel"
                                                              onClick={() => onRemoveAttachmentWarning("wordInfo")}
                                                            >
                                                              X
                                                            </span>
                                                          )}
                                                      </span>
                                                    </div>
                                                  )}
                                                  {(noteData?.noteSecretaryDTO?.some(
                                                    (obj) =>
                                                      obj.secretaryEmail === accounts[0].username &&
                                                      obj.approverEmail === noteData?.currentActioner) ||
                                                    (isGistSecretary && noteData?.currentActionerDesignation === MDNameTitle)
                                                  ) && (
                                                      <div className="_fileFormatHintMsg">
                                                        Allowed Formats (pdf only).Up to 5MB max.
                                                      </div>
                                                    )}
                                                </div>
                                              </div>
                                            </>
                                          )}


                                        {item.id === "Comments & Workflow Log" && item.isVisible && (
                                          <div className="table-responsive">
                                            <table className="ib-forms-custom-table tableStyle mob_table">
                                              <tbody>
                                                <tr>
                                                  <th className="approvalform-tableCol-width-ActionBy">Name</th>
                                                  <th className="approvalform-tableCol-width-action">Action</th>
                                                  <th className="approvalform-tableCol-width-Comments">Comments</th>
                                                  <th className="approvalform-tableCol-width-actiondate">Action Date</th>
                                                </tr>

                                                {(() => {
                                                  const shownPdfComments = new Set();

                                                  return noteData.commentAuditLogDto?.map((comment, index) => {
                                                    return (
                                                      <tr key={index}>
                                                        <td className="approvalform-tableCol-width-ActionBy">
                                                          {comment.name}
                                                        </td>

                                                        <td className="approvalform-tableCol-width-action">
                                                          <span>
                                                            {comment.action}
                                                            <span
                                                              className="k-icon k-font-icon k-i-information allIconsforPrimary-btn"
                                                              style={{ cursor: "pointer", marginLeft: "8px" }}
                                                              title="View Details"
                                                              onClick={() => handleIconClick(comment)}
                                                            />
                                                          </span>
                                                        </td>

                                                        <td className="approvalform-tableCol-width-Comments">
                                                          {(() => {
                                                            const displayComment = (comment?.comments || "").trim();

                                                            if (!displayComment || displayComment === "NA") {
                                                              return "NA";
                                                            }

                                                            const entries = displayComment
                                                              .split(/\||\r?\n/)
                                                              .map((entry) => entry.trim())
                                                              .filter(Boolean)
                                                              .filter((entry) => !/\[PDF Highlight\]:/i.test(entry));

                                                            const finalEntries = [];

                                                            entries.forEach((entry) => {
                                                              const isPdfComment = /\[PDF Comment\]:/i.test(entry);

                                                              if (isPdfComment) {
                                                                const duplicateKey = entry
                                                                  .replace(/\s+/g, " ")
                                                                  .replace(/\[PDF Comment\]:/i, "")
                                                                  .trim()
                                                                  .toLowerCase();

                                                                if (shownPdfComments.has(duplicateKey)) return;

                                                                shownPdfComments.add(duplicateKey);
                                                              }

                                                              finalEntries.push(entry);
                                                            });

                                                            if (finalEntries.length === 0) {
                                                              return "NA";
                                                            }

                                                            return finalEntries.map((entry, i) => {
                                                              if (/^\d+,\s*[^,]+,\s*\[PDF Comment\]:/i.test(entry)) {
                                                                const parts = entry.split(",");
                                                                const pageNo = parts[0]?.trim() || "NA";
                                                                const lineNo = parts[1]?.trim() || "NA";

                                                                const userComment = entry
                                                                  .replace(/^\d+,\s*[^,]+,\s*\[PDF Comment\]:/i, "")
                                                                  .trim();

                                                                return (
                                                                  <div key={i}>
                                                                    PageNo - {pageNo} , Line No - {lineNo} , UserComment - {userComment}
                                                                  </div>
                                                                );
                                                              }

                                                              return (
                                                                <div key={i}>
                                                                  {entry.replace(/\[PDF Comment\]:\s*/gi, "").trim()}
                                                                </div>
                                                              );
                                                            });
                                                          })()}
                                                        </td>

                                                        <td className="approvalform-tableCol-width-actiondate">
                                                          {getdateconversion(comment.actionDate)}
                                                        </td>
                                                      </tr>
                                                    );
                                                  });
                                                })()}
                                              </tbody>
                                            </table>
                                          </div>
                                        )}
                                        {item.id === "File Attachments" &&
                                          item.isVisible && (
                                            <div>
                                              {noteData &&
                                                noteData?.notePdfPath !== null && (
                                                  <p>
                                                    <b>Main Note link :</b>
                                                    {/* 04/05 -Removed link and added Anchor tag */}
                                                    <span
                                                      className="link-style"
                                                      onClick={() => downloadBase64PDFFile(noteData?.notePdfPath)}
                                                    >
                                                      <b>{noteData?.notePdfFileName}</b>
                                                    </span>
                                                  </p>
                                                )}
                                              {(isSecretary ||
                                                getNoteSecretaryUserInfo?.some(
                                                  (obj) =>
                                                    obj.approverEmail ===
                                                    accounts[0].username &&
                                                    obj.approverType === 2
                                                )) && (
                                                  <>
                                                    {noteData &&
                                                      noteData?.noteWordPath !==
                                                      null && (
                                                        <p>
                                                          <b>Main Note word document link :</b>
                                                          <span
                                                            className="link-style"
                                                            onClick={() =>
                                                              downloadBase64WordFile(
                                                                noteData?.noteWordPath,
                                                                noteData?.noteWordFileName
                                                              )
                                                            }
                                                          >
                                                            <b>{noteData?.noteWordFileName}</b>
                                                          </span>
                                                        </p>
                                                      )}
                                                  </>
                                                )}

                                              <p><b>Support Documents :</b></p>
                                              <div className="table-responsive">
                                                <table className="tableStyle">
                                                  <tbody>
                                                    {" "}
                                                    <tr>
                                                      <th className="approvalform-tableCol-width-4">
                                                        Document link
                                                      </th>
                                                      <th className="approvalform-tableCol-width-3">
                                                        Attached By
                                                      </th>
                                                      <th className="approvalform-tableCol-width-3">
                                                        Attached Date
                                                      </th>
                                                    </tr>
                                                    {noteData?.noteSupportingDocumentsDTO?.map(
                                                      (doc) => (
                                                        <tr
                                                          key={
                                                            doc.noteSupportingDocumentId
                                                          }
                                                        >
                                                          <td className="approvalform-tableCol-width-4">
                                                            <span
                                                              className="link-style"
                                                              onClick={() =>
                                                                getSupportDocHyperlink(
                                                                  doc.supportingDocumentPath,
                                                                  doc.supportingDocumentFileName
                                                                )
                                                              }
                                                            >
                                                              {
                                                                doc.supportingDocumentFileName
                                                              }
                                                            </span>
                                                          </td>
                                                          {/* Bug fix - 293 - 27/03 */}
                                                          <td className="approvalform-tableCol-width-3">
                                                            {doc.createdByName}
                                                          </td>
                                                          <td className="approvalform-tableCol-width-3">
                                                            {getdateconversion(
                                                              doc.createdDate
                                                            )}
                                                          </td>
                                                        </tr>
                                                      )
                                                    )}
                                                  </tbody>
                                                </table>
                                              </div>
                                            </div>
                                          )}
                                        {item.id === "Mark for Information Section" &&
                                          item.isVisible && (
                                            <>
                                              <div className="SectionRow row">
                                                <div>
                                                  <div
                                                    onKeyDown={(e) =>
                                                      e.stopPropagation()
                                                    }
                                                    className="_markInfocontainer"
                                                  >
                                                    <MultiColumnComboBox
                                                      data={orgEmployees}
                                                      filterable={true}
                                                      columns={
                                                        isMobile
                                                          ? mobileColumns
                                                          : orgUsersPplPickerColumnMapping
                                                      }
                                                      value={
                                                        selectedMarkUserInfo === null
                                                          ? null
                                                          : selectedMarkUserInfo.displayName
                                                      }
                                                      onFilterChange={
                                                        onFillterALLUser
                                                      }
                                                      onChange={onchangehanleMarkInfo}
                                                      className="_markinfoCombobox"
                                                      placeholder="Add Users..."
                                                    />
                                                    <Button
                                                      onClick={addUserInForMarkInfo}
                                                      className="formBtnColor non_mobileResponsive"
                                                    >
                                                      <span className="k-icon k-font-icon k-i-plus cursor allIconsforPrimary-btn"></span>
                                                      Add User
                                                    </Button>
                                                    <Button
                                                      onClick={addUserInForMarkInfo}
                                                      className="formBtnColor _mobileResponsive "
                                                    >
                                                      <span className="k-icon k-font-icon k-i-plus cursor allIconsforPrimary-btn"></span>
                                                    </Button>
                                                  </div>

                                                  <div className="table-responsive">
                                                    <table className="tableStyle">
                                                      <thead>
                                                        <tr>
                                                          <th className="approvalform-tableCol-width-1">
                                                            S.No
                                                          </th>
                                                          <th className="approvalform-tableCol-width-6">
                                                            User Info
                                                          </th>
                                                          <th className="approvalform-tableCol-width-3">
                                                            Action
                                                          </th>
                                                        </tr>
                                                      </thead>
                                                      <tbody>
                                                        {markInfoUserEmail?.map(
                                                          (obj, index) => (
                                                            <tr key={index}>
                                                              <td className="approvalform-tableCol-width-1">
                                                                {index + 1}
                                                              </td>
                                                              {/* Bug fix - 293 - 27/03 */}
                                                              <td className="approvalform-tableCol-width-6">
                                                                {/* Bug - 293 29/03 */}
                                                                {obj.markedEmailName}
                                                              </td>
                                                              <td className="approvalform-tableCol-width-3">
                                                                {/* _mobileResponsive */}
                                                                <Button
                                                                  className="non_mobileResponsive"
                                                                  onClick={() =>
                                                                    removeUserFromMarkInfo(
                                                                      index
                                                                    )
                                                                  }
                                                                >
                                                                  <span className="k-icon k-font-icon k-i-delete k-i-trash cursor allIconsforPrimary-btn"></span>
                                                                  Delete
                                                                </Button>
                                                                <Button
                                                                  className="_mobileResponsive"
                                                                  onClick={() =>
                                                                    removeUserFromMarkInfo(
                                                                      index
                                                                    )
                                                                  }
                                                                >
                                                                  <span className="k-icon k-font-icon k-i-delete k-i-trash cursor allIconsforPrimary-btn"></span>
                                                                </Button>
                                                              </td>
                                                            </tr>
                                                          )
                                                        )}
                                                      </tbody>
                                                    </table>
                                                  </div>
                                                </div>
                                              </div>
                                              {markInfoUserEmail?.length > 0 && (
                                                <Button
                                                  themeColor="info"
                                                  className="formBtnColor"
                                                  onClick={addmarkInfotoDataBase}
                                                >
                                                  <span className="k-icon k-font-icon k-i-launch cursor allIconsforPrimary-btn"></span>
                                                  Submit
                                                </Button>
                                              )}
                                            </>
                                          )}
                                      </ExpansionPanelContent>
                                    )}
                                  </>
                                )}
                              </Reveal>
                            </ExpansionPanel>
                          ) : null
                        )}
                      </div>
                      <div
                        className={
                          isPDFFullWidth ? "homesectionPdf-2" : "viewFormSection-2"
                        }

                      >
                        {/* // ========== PDF ANNOTATION CHANGE START ==========
// MODIFIED - Added proper zoom handling when entering/exiting full screen */}
                        {isPDFFullWidth ? (
                          <span
                            className="cursor pdfHideandShowIcons"
                            id='zoomFill'
                            onClick={() => {
                              setIsPDFFullWidth(false);
                              setIsResolutionChecked(true);

                              requestAnimationFrame(() => {
                                updatePanelHeight();

                                requestAnimationFrame(() => {
                                  if (pdfAnnotatorRef.current?.setZoom) {
                                    pdfAnnotatorRef.current.setZoom(DEFAULT_PDF_ZOOM);
                                  }
                                });
                              });
                            }}
                          >
                            <span className="k-icon k-font-icon k-i-fullscreen-exit k-i-full-screen-exit"></span>
                            <div className="fullScreen">Exit<br></br>Full<br></br>Scr</div>
                          </span>
                        ) : (
                          <span
                            className="cursor pdfHideandShowIcons"
                            id='zoomExit'
                            onClick={() => {
                              setIsPDFFullWidth(true);
                              setIsResolutionChecked(true);

                              requestAnimationFrame(() => {
                                updatePanelHeight();

                                requestAnimationFrame(() => {
                                  if (pdfAnnotatorRef.current?.setZoom) {
                                    pdfAnnotatorRef.current.setZoom(FULLSCREEN_PDF_ZOOM);
                                  }
                                });
                              });
                            }}
                          >
                            <span className="k-icon k-font-icon k-i-fullscreen k-i-fullscreen-enter"></span>
                            <div className="fullScreen">Full<br></br>Scr</div>
                          </span>
                        )}
                        {/* 
                        // ========== PDF ANNOTATION CHANGE END ========== */}

                        <div
                          className="pdf-viewer _pdfViewer"
                          ref={pdfViewerRef}
                          style={{
                            height: isPDFFullWidth ? "90vh" : `calc(${panelHeight}px)`,
                            // overflow:  "auto !important",
                            border: isPDFFullWidth ? "none" : "1px solid #00000030",
                            position: "relative"
                          }}
                        >




                          {/* 

// ========== PDF ANNOTATION CHANGE START ==========
// MODIFIED - Replaced ReactPDFViewerApp with PdfAnnotator component
// Old code: <ReactPDFViewerApp pdfBlob={pdfBlob} fileName={noteData?.noteNumber} />
// New code: PdfAnnotator with full annotation support */}
                          {pdfLoading ? (
                            <div className="pdf-loading-container" style={{ padding: "2rem", textAlign: "center" }}>
                              <div style={{ marginBottom: "0.5rem" }}>Loading PDF... {pdfProgress}%</div>
                              <div
                                style={{
                                  width: "80%",
                                  margin: "0 auto",
                                  height: "8px",
                                  backgroundColor: "#e0e0e0",
                                  borderRadius: "4px",
                                  overflow: "hidden"
                                }}
                              >
                                <div
                                  style={{
                                    width: `${pdfProgress}%`,
                                    height: "100%",
                                    backgroundColor: "#007bff",
                                    transition: "width 0.3s ease"
                                  }}
                                />
                              </div>
                            </div>
                          ) : (
                            // Render PDF Pages
                            // <ReactPDFViewerApp pdfBlob={pdfBlob} fileName={noteData?.noteNumber} />
                            // In your ENoteViewForm component, replace the existing PDF viewer:



                            //priti shelke - 08-04-2026
                            <PdfAnnotator
                              ref={pdfAnnotatorRef}
                              pdfBlob={pdfBlob}
                              fileName={noteData?.notePdfFileName || noteData?.noteNumber || "document.pdf"}
                              currentUser={accounts?.[0]?.username || "User"}
                              isReadOnly={false}
                              generalDetails={noteData}
                              initialZoom={Number(zoomLevel) || 1}
                              isFullScreen={isPDFFullWidth}
                              initialAnnotations={pdfInitialAnnotations}
                              draftStorageKey={getPdfDraftStorageKey(noteData?.noteId || id?.id, accounts?.[0]?.username)}
                              onAnnotationsChange={(allAnnotations) => {
                                saveCurrentUserDraftAnnotations(allAnnotations);
                              }}
                              onAddComment={(comment) => {
                                setPdfAnnotationComments((prev) => {
                                  const exists = prev.some(
                                    (x) =>
                                      x.pdfAnnotationId === comment.id ||
                                      Number(x.annotationId) === Number(comment.annotationId)
                                  );
                                  if (exists) return prev;

                                  return [
                                    {
                                      PageNo: comment.page,
                                      // ✅ DocRef मध्ये selected text नाही, line number ठेवायचा
                                      DocRef: String(
                                        comment.lineNumberRange ||
                                        comment.lineNumber ||
                                        "NA"
                                      ),
                                      Comments: `[PDF Comment]: ${comment.text || ""}`,
                                      isEdit: false,
                                      fromPDF: true,
                                      annotationType: "comment",
                                      pdfAnnotationId: comment.id,
                                      annotationId:
                                        comment.annotationId ||
                                        comment?.serverResponse?.annotationId ||
                                        null,
                                      lineNumber:
                                        comment.lineNumberRange ||
                                        comment.lineNumber ||
                                        "NA",
                                      selectedText: comment.selectedText || "",
                                      timestamp: comment.timestamp,
                                      user: comment.user,
                                      normalizedRects: comment.normalizedRects || [],
                                    },
                                    ...prev,
                                  ];
                                });
                              }}

                              onUpdateComment={(updatedComment) => {
                                setPdfAnnotationComments((prev) =>
                                  prev.map((ann) =>
                                    ann.pdfAnnotationId === updatedComment.id ||
                                      Number(ann.annotationId) === Number(updatedComment.annotationId)
                                      ? {
                                        ...ann,
                                        Comments: `[PDF Comment]: ${updatedComment.text || ""}`,
                                        // ✅ DocRef मध्ये line number
                                        DocRef: String(
                                          updatedComment.lineNumberRange ||
                                          updatedComment.lineNumber ||
                                          ann.lineNumber ||
                                          "NA"
                                        ),
                                        PageNo: updatedComment.page,
                                        annotationType: "comment",
                                        annotationId:
                                          updatedComment.annotationId ||
                                          updatedComment?.serverResponse?.annotationId ||
                                          ann.annotationId ||
                                          null,
                                        lineNumber:
                                          updatedComment.lineNumberRange ||
                                          updatedComment.lineNumber ||
                                          ann.lineNumber ||
                                          "NA",
                                        selectedText: updatedComment.selectedText || ann.selectedText || "",
                                        timestamp: updatedComment.timestamp,
                                        user: updatedComment.user,
                                        normalizedRects:
                                          updatedComment.normalizedRects ||
                                          ann.normalizedRects ||
                                          [],
                                      }
                                      : ann
                                  )
                                );
                              }}

                              onAddHighlight={(highlight) => {
                                // ✅ Important:
                                // Highlight already saved in AnnotationDetails from PdfAnnotator.
                                // Do NOT add highlight into pdfAnnotationComments.
                                // Because pdfAnnotationComments is used for NoteApproverCommentsDTO on approval.
                                console.log("Highlight saved only in AnnotationDetails. Not adding to NoteApproverComments payload.", highlight);
                              }}
                              onResetAnnotations={() => {
                                setPdfAnnotationComments([]);
                                clearCurrentUserDraftAnnotations();
                              }}
                            />
                            // ========== PDF ANNOTATION CHANGE END ==========

                            //priti shelke - 08-04-2026




                          )}
                        </div>

                      </div>

                    </div>


                  </div>
                </fieldset>

              </div>
            </div>
            {/* {isLoading && <PageLoader />} */}
            <div className="approvalAction-foreNoteForm-contaioner">
              <div>
                {approverChecking() && (
                  <span className="eNote-ApprovalButton">
                    <Button onClick={approvalFunction} themeColor="success" disabled={actionInProgress}>
                      <span className="k-icon k-font-icon  k-i-track-changes-accept cursor allIconsforPrimary-btn"></span>
                      {isApprover() ? "Noted" : "Approve"}
                      {/* Approve */}
                    </Button>
                  </span>
                )}
                {approverChecking() && (
                  <span className="eNote-ApprovalButton">
                    <Button onClick={rejectFunction} themeColor="error" disabled={actionInProgress}>
                      <span className="k-icon k-font-icon  k-i-track-changes-reject cursor allIconsforPrimary-btn"></span>
                      Reject
                    </Button>
                  </span>
                )}
                {approverChecking() && (
                  <span className="eNote-ApprovalButton">
                    <Button onClick={referFunction} className="formBtnColor" disabled={actionInProgress}>
                      <span className="k-icon k-font-icon k-icon-xl k-i-arrow-root cursor allIconsforPrimary-btn"></span>
                      Refer
                    </Button>
                  </span>
                )}
                {noteData?.currentActioner === accounts[0].username &&
                  isreferredUser &&
                  noteData.status === 9 &&
                  noteData !== null && (
                    <span className="eNote-ApprovalButton">
                      <Button onClick={referBackfunction} className="formBtnColor" disabled={actionInProgress}>
                        <span className="k-icon k-font-icon  k-i-arrow-drill cursor allIconsforPrimary-btn"></span>
                        Refer Back
                      </Button>
                    </span>
                  )}
              </div>
              <div>
                {approverChecking() && (
                  <span className="eNote-ApprovalButton">
                    <Button onClick={returnFunction} className="formBtnColor" disabled={actionInProgress}>
                      <span className="k-icon k-font-icon  k-i-undo cursor allIconsforPrimary-btn"></span>
                      Return
                    </Button>
                  </span>
                )}
                {noteData &&
                  noteData.status === 2 &&
                  noteData.createdBy === accounts[0].username &&
                  noteData.noteApproversDTO?.every(
                    (obj) => obj.approverStatus === 1 || obj.approverStatus === 2
                  ) && (
                    <span className="eNote-ApprovalButton">
                      <Button className="formBtnColor" onClick={handleCallBack} disabled={actionInProgress}>
                        <span className="k-icon k-font-icon  k-i-caret-alt-to-left cursor allIconsforPrimary-btn"></span>
                        Call Back
                      </Button>
                    </span>
                  )}
                {noteData &&
                  (noteData.status === 3 || noteData.status === 9) &&
                  noteData.createdBy === accounts[0].username && (
                    <span className="eNote-ApprovalButton">
                      <Button
                        className="formBtnColor"
                        onClick={onClickChangeApprover} disabled={actionInProgress}
                      >
                        <span className="k-icon k-font-icon  	k-i-user cursor allIconsforPrimary-btn"></span>
                        Change Approver
                      </Button>
                    </span>
                  )}

                {/* change added submit button in bottom of the page  -16/04 */}
                {/* {
                  (noteData?.noteSecretaryDTO?.some(
                    (obj) =>
                      obj.secretaryEmail === accounts[0].username &&
                      obj.approverEmail === noteData?.currentActioner
                  ) || (isGistSecretary && noteData?.currentActionerDesignation === MDNameTitle)) && ( */}
                {(noteData?.noteSecretaryDTO?.some(
                  (obj) => obj.secretaryEmail === accounts[0].username &&
                    obj.approverEmail === noteData?.currentActioner
                ) && (isGistSecretary && noteData?.currentActionerDesignation === MDNameTitle)) && (
                    <span className="eNote-ApprovalButton">
                      <Button onClick={onSubmitForAttcGISTDoc} className="formBtnColor" disabled={actionInProgress}>
                        <span className="k-icon k-font-icon k-i-launch cursor allIconsforPrimary-btn"></span>
                        <span>Submit</span>
                      </Button>
                    </span>
                  )
                }
                <span className="eNote-ApprovalButton">
                  <Button onClick={redirectHomePage}>
                    <span className="k-icon k-font-icon  k-i-x-circle cursor allIconsforPrimary-btn"></span>
                    Exit
                  </Button>
                </span>
              </div>
            </div>
            <Footer /></>
        )
        }
      </>

      {dialogOpen && (
        <Dialog
          title={<CustomConfirmationTitleBar />}
          onClose={toggleDailogForConfirmation}
        >
          <p className="dialogcontent_">{confirmDailogObj.Confirmtext}</p>
          <p className="dialogcontent_">{confirmDailogObj.Description}</p>

          {isCurrentATRCreator &&
            selectedAssigneeUsersinfo?.length === 0 &&
            actionBtn === "Approve" && (
              <p className="dialogcontent_">
                Note: ATR assignees are not found. So for this approval request
                ATR assigness will not create.
              </p>
            )}

          <DialogActionsBar>
            <Button className="formBtnColor" onClick={onCofiormation}>
              <span className="k-icon k-font-icon  k-i-checkmark-circle cursor allIconsforPrimary-btn"></span>
              Confirm
            </Button>

            <Button onClick={toggleDailogForConfirmation}>
              <span className="k-icon k-font-icon  k-i-close-circle cursor allIconsforPrimary-btn"></span>
              Cancel
            </Button>
          </DialogActionsBar>
        </Dialog>
      )}

      {
        dialogSuccess && (
          <Dialog title={<CustomSuccussTitleBar />} onClose={redirectHomePage}>
            <p className="dialogcontent_">{successMsg}</p>
            <DialogActionsBar>
              <Button className="notifyDailogOkBtn" onClick={redirectHomePage}>
                <span className="k-icon k-font-icon  k-i-redo cursor allIconsforPrimary-btn"></span>
                Ok
              </Button>
            </DialogActionsBar>
          </Dialog>
        )
      }
      {
        dialogFailure && (
          <Dialog
            title={<CustomSuccussTitleBar />}
            onClose={() => setDialogFailure(false)}
          >
            <p className="dialogcontent_">
              {`Something went wrong: ${statusMessage}`}
            </p>
            <DialogActionsBar>
              <Button
                className="notifyDailogOkBtn"
                onClick={() => setDialogFailure(false)}
              >
                <span className="k-icon k-font-icon  k-i-redo cursor allIconsforPrimary-btn"></span>
                Ok
              </Button>
            </DialogActionsBar>
          </Dialog>
        )
      }
      {
        callbackvisible && (
          <Dialog
            title={<CustomSuccussTitleBar />}
            onClose={() => setCallBackVisible(false)}
          >
            <p className="dialogcontent_">
              The request for call back has been successfull.
            </p>
            <DialogActionsBar>
              <Button onClick={handleCloseDialog} className="notifyDailogOkBtn">
                <span className="k-icon k-font-icon  k-i-redo cursor allIconsforPrimary-btn"></span>
                Ok
              </Button>
            </DialogActionsBar>
          </Dialog>
        )
      }
      {
        changeapprovervisible && (
          <Dialog
            title={<CustomSuccussTitleBar />}
            onClose={() => setChangeApproverVisible(false)}
          >
            <p className="dialogcontent_">
              The current actioner(Approver/Reviewer/Referee) has been updated
              successfully.
            </p>
            <DialogActionsBar>
              <Button onClick={handleCloseDialog} className="notifyDailogOkBtn">
                <span className="k-icon k-font-icon  k-i-redo cursor allIconsforPrimary-btn"></span>
                Ok
              </Button>
            </DialogActionsBar>
          </Dialog>
        )
      }
      {/* dialog for validation */}
      {
        userComboValidationDialog && (
          <Dialog
            title={<CustomSuccussTitleBar />}
            onClose={() => { setUserComboValidationDialog(false); setActionInProgress(false); }}
          >
            <p className="dialogcontent_">{userNotifyInfo}</p>
            {/* change 16/04 gist doc validation */}
            {actionBtn === "AttchGistDoc" && (
              <p className="dialogcontent_">
                <strong>Note:</strong> Invalid files are not allowed.
              </p>
            )}

            <DialogActionsBar>
              <Button onClick={oncloseUseNotify} className="notifyDailogOkBtn">
                <span className="k-icon k-font-icon  k-i-redo cursor allIconsforPrimary-btn"></span>
                Ok
              </Button>
            </DialogActionsBar>
          </Dialog>
        )
      }
      {/* add refer and change approver */}
      {
        dialogForRefereeAndChangeApprover && (
          <Dialog
            title={<CustomTitleBar />}
            onClose={() => {
              setDialogForRefereeAndChangeApprover(false);
              setActionInProgress(false);
            }}
            className="dialogcontent_Refer"
          >
            {actionBtn === "Refer" && (
              <>
                <label className="addReferee-label-ChangeApprover-label">
                  Add Referee
                  <br></br>
                  <MultiColumnComboBox
                    data={orgEmployees}
                    filterable={true}
                    columns={
                      isMobile ? mobileColumns : orgUsersPplPickerColumnMapping
                    }
                    value={
                      selectedrefer === null ? null : selectedrefer.displayName
                    }
                    onFilterChange={onFillterALLUser}
                    onChange={selectedReferrUser}
                    style={{ width: isMobile ? "100%" : "500px" }} // Adjust width
                    placeholder="Add Referee..."
                  />
                </label>
                <br />
                <label className="addReferee-label-ChangeApprover-label">
                  Comments
                  <br></br>
                  <textarea
                    name="Comments"
                    className="general-comments-Comments"
                    row={1}
                    value={referCommentsObj.Comments}
                    onChange={handelReferGenralComments}
                  ></textarea>
                </label>
              </>
            )}
            {actionBtn === "ChangeApprover" && (
              <>
                <label className="addReferee-label-ChangeApprover-label">
                  {" "}
                  Change approver
                  <span className="required-asterisk">*</span>
                  <MultiColumnComboBox
                    data={orgEmployees}
                    filterable={true}
                    columns={
                      isMobile ? mobileColumns : orgUsersPplPickerColumnMapping
                    }
                    value={
                      changeApprovercombovalue === null
                        ? null
                        : changeApprovercombovalue.displayName
                    }
                    onFilterChange={onFillterALLUser}
                    onChange={handleChangeApproverCombo}
                    placeholder="Change Approver..."
                  />
                </label>
                <br />
                <br />
                {changeApproverhaveSecretary && (
                  <>
                    <label className="addReferee-label-ChangeApprover-label">
                      {" "}
                      Word Document
                      <span className="required-asterisk">*</span>
                    </label>
                    <div className="Attachemntfileinfo-ind addReferee-label-ChangeApprover-label">
                      <input
                        type="file"
                        id="noteWordDocfile"
                        key={fileKey} // Unique key to force re-render
                        // onChange={(event) => noteWordconvertBase64Word(event, 'WORD')} // WORD as file category
                        onChange={noteWordconvertBase64Word}
                        className="_ConvertPDFToBase64"
                      />
                    </div>
                    {noteWordDoc?.noteWordDocInfo.fileName && (
                      <div className="Attachemntfileinfo-ind addReferee-label-ChangeApprover-label">
                        <span className="attachemntIconInfoConationer">
                          <span className="AttchemntIconWraper">
                            <SvgIcon
                              icon={getFileIcon(
                                noteWordDoc?.noteWordDocInfo.fileName
                              )}
                              size="xxlarge"
                            />
                            <span className="attachemnrt-warningifoConatiner">
                              <div className="attachemnrt-warningifo-fileinfo">
                                {noteWordDoc?.noteWordDocInfo.fileName}
                              </div>
                              <span className="inCorrectFileError">
                                {noteWordDoc?.noteWordDocInfo.warningMsg}
                              </span>
                            </span>
                          </span>
                          <span
                            className="AttchemntIconWraperCancel"
                            onClick={onNoteWordDocRemoveAttachmentWarning}
                          >
                            X
                          </span>
                        </span>
                      </div>
                    )}
                    <div className="addReferee-label-ChangeApprover-label _fileFormatHintMsg">
                      Allowed only one Word Document. Upto 10MB max.
                    </div>
                  </>
                )}
              </>
            )}
            <DialogActionsBar>
              <Button
                className="formBtnColor"
                onClick={onSubmitforAddReferandChangeApprover}
              >
                <span className="k-icon k-font-icon  k-i-launch cursor allIconsforPrimary-btn"></span>
                Submit
              </Button>
              <Button onClick={() => {
                setDialogForRefereeAndChangeApprover(false);
                setActionInProgress(false);
              }}>
                <span className="k-icon k-font-icon  k-i-close-circle cursor allIconsforPrimary-btn"></span>
                Cancel
              </Button>
            </DialogActionsBar>
          </Dialog>
        )
      }

      {/* passcode verification dialog*/}
      {
        passcodeVerification && (
          <Dialog title="Passcode Verification" className="mobile_passcode" onClose={handlepasscodeClose}>
            <form className="form-container dialogcontent_ ">
              <label>
                Enter your passcode for verification:
                <div className="passcode-input-wrapper">
                  <input
                    type={isPasscodeVisible ? "text" : "password"}
                    value={passcode}
                    onChange={handlePasscodeChange}
                    className="passcode-input"
                    maxLength="6"
                    pattern="\d*"
                    title="Please enter 6-characters, Combination of Alphabets and Numbers"
                  />
                  <span
                    className="eye-icon-view"
                    tabIndex="0" // Makes the span focusable for keyboard events
                    onMouseDown={() => setIsPasscodeVisible(true)}    // Show on mouse down
                    onMouseUp={() => setIsPasscodeVisible(false)}     // Hide on mouse up
                    onKeyDown={(e) => handleKeyDown(e)}            // Show on key down (Enter or Space)
                    onKeyUp={(e) => handleKeyUp(e)}
                    onTouchStart={() => setIsPasscodeVisible(true)}   // Show on touch start (mobile)
                    onTouchEnd={() => setIsPasscodeVisible(false)}
                  >
                    {isPasscodeVisible ? (
                      <img alt="view" src={view} />
                    ) : (
                      <img alt="hide" src={hide} />
                    )}
                  </span>
                </div>
              </label>
              <div>{error && <span className="error">{error}</span>}</div>
            </form>
            <DialogActionsBar>
              <Button
                className="formBtnColor"
                onClick={passcodeVerificationFunction}
              >
                <span className="k-icon k-font-icon  k-i-launch cursor allIconsforPrimary-btn"></span>
                <span>Verify</span>
              </Button>
              <Button onClick={handlepasscodeClose}>
                <span className="k-icon k-font-icon  k-i-close-circle cursor allIconsforPrimary-btn"></span>
                <span>Cancel</span>
              </Button>
            </DialogActionsBar>
          </Dialog>
        )
      }

      {mobileCommentsDialog && (
        <Dialog title="Add Comments" onClose={handleCloseCommentsDialog}>
          <form className="form-container dialogcontent_Comments">
            <div className="row">
              <div className="col-sm-12">
                <div onKeyDown={(e) => e.stopPropagation()}>
                  <label>Page No: </label>
                  <br></br>
                  <input
                    className="general-comments-pageNo"
                    type="text"
                    name="PageNo"
                    value={generalcmtInfoobj.PageNo}
                    onChange={handelGenralComments}
                  />
                </div>
              </div>
              <div className="col-sm-12">
                <div onKeyDown={(e) => e.stopPropagation()}>
                  <label>Doc Reference: </label>
                  <br></br>
                  <input
                    className="general-comments-DocRef"
                    name="DocRef"
                    value={generalcmtInfoobj.DocRef}
                    onChange={handelGenralComments}
                  />
                </div>
              </div>
              <div className="col-sm-12">
                <div onKeyDown={(e) => e.stopPropagation()}>
                  <label>
                    Comments: <span className="required-asterisk">*</span>
                  </label>
                  <textarea
                    name="Comments"
                    maxLength={max}
                    className="general-comments-Comments"
                    row={5}
                    value={generalcmtInfoobj.Comments}
                    onChange={handelGenralComments}
                  ></textarea>
                  <Hint direction={"end"}>
                    {generalcmtInfoobj.Comments.length} / {max}
                  </Hint>
                  {errorMessageEdit && (
                    <div className="inCorrectFileError">
                      {errorMessageEdit}
                    </div>
                  )}
                  <Button
                    className="formBtnColor mt-2 button_center"
                    onClick={handelAddCmts}
                  >
                    Save
                  </Button>
                </div>
              </div>
            </div>
          </form>
        </Dialog>
      )
      }

      {mobileCommentsEditDialog && (
        <Dialog title="Edit Comments" onClose={handleCloseCommentsDialog}>
          <form className="form-container dialogcontent_Comments">
            <div
              onKeyDown={(e) => e.stopPropagation()}
              className="table"
            >
              {generalcmtforAllCmt?.map((obj, index) => (
                <div key={index}>
                  {obj.isEdit ? (
                    <div onKeyDown={(e) => e.stopPropagation()}>
                      <label className="text-left">Page No:</label>
                      <input
                        type="text"
                        className="general-comments-pageNo"
                        value={obj.PageNo}
                        name="PageNo"
                        id={index}
                        onChange={handelGenralCommentsEdit}
                      />
                    </div>
                  ) : (
                    obj.PageNo
                  )}
                  {obj.isEdit ? (
                    <div onKeyDown={(e) => e.stopPropagation()}>
                      <label>Doc Reference: </label>
                      <input
                        type="text"
                        className="general-comments-DocRef"
                        value={obj.DocRef}
                        name="DocRef"
                        id={index}
                        onChange={handelGenralCommentsEdit}
                      />
                    </div>
                  ) : (
                    obj.DocRef
                  )}
                  {obj.isEdit ? (
                    <div onKeyDown={(e) => e.stopPropagation()}>
                      <label>
                        Comments: <span className="required-asterisk">*</span>
                      </label>
                      <textarea
                        value={obj.Comments}
                        maxLength={max}
                        className="general-comments-Comments"
                        name="Comments"
                        row={5}
                        id={index}
                        onChange={handelGenralCommentsEdit}
                      ></textarea>
                      <Hint direction={"end"}>
                        {obj.Comments.length} / {max}
                      </Hint>
                    </div>
                  ) : (
                    obj.Comments
                  )}
                </div>
              ))}
              {errorMessageEdit && (
                <div className="inCorrectFileError">
                  {errorMessageEdit}
                </div>
              )}
              <Button
                className="formBtnColor mt-2 button_center" onClick={handleMobileSave} // No need to pass index since it's handled in state
              >
                Update
              </Button>
            </div>
          </form>
        </Dialog>
      )
      }

      {/* Valid passcode dialog*/}
      {
        validPasscode && (
          <Dialog
            title="Passcode Verification"
            onClose={() => setValidPasscode(false)}
          >
            <div className="dialogcontent_">
              <p>{validMsg}</p>
            </div>
            <DialogActionsBar>
              <Button className="formBtnColor" onClick={handleRedirectToPasscode}>
                <span className="k-icon k-font-icon  k-i-launch cursor allIconsforPrimary-btn"></span>
                Create Passcode
              </Button>
              <Button onClick={() => setValidPasscode(false)}>
                <span className="k-icon k-font-icon  k-i-close-circle cursor allIconsforPrimary-btn"></span>
                Cancel
              </Button>
            </DialogActionsBar>
          </Dialog>
        )
      }
      {gistDocPopup && approverChecking() &&
        (
          <Dialog
            title="Gist Document"
            onClose={() => setGistDocPopup(false)}
            className="gist_dialogcontent_content"
          >
            <div className="gist_dialogcontent_">
              <ReactPDFViewerGist
                pdfBlob={gistPdfBlob}
                fileName={noteData?.noteNumber}
              />
            </div>

            <DialogActionsBar>
              <Button onClick={() => setGistDocPopup(false)} className="dialogCloseGist">
                <span className="k-icon k-font-icon  k-i-close-circle cursor allIconsforPrimary-btn"></span>
                Close
              </Button>
            </DialogActionsBar>
          </Dialog>
        )
      }
      {
        deleteNotification && (
          <Dialog
            title={<CustomSuccussTitleBar />}
            onClose={() => setDeleteNotification(false)}
          >
            <div className="dialogAlignment">
              <p>Do you want to delete this comment?</p>
            </div>
            <DialogActionsBar>
              <Button className="formBtnColor" onClick={handleDeleteConfirmation}>
                <span className="k-icon k-font-icon k-i-launch cursor allIconsforPrimary-btn"></span>
                <span>Yes</span>
              </Button>
              <Button onClick={() => setDeleteNotification(false)}>
                <span className="k-icon k-font-icon k-i-close-circle cursor allIconsforPrimary-btn"></span>
                <span>No</span>
              </Button>
            </DialogActionsBar>
          </Dialog>
        )
      }


      {isDialogOpen &&
        selectedComment &&
        (
          <Dialog title={"Individual TAT Details"} onClose={() => setIsDialogOpen(false)}>
            <div className="dialogcontent_">
              <table className="ib-forms-custom-table tableStyle">
                <tbody>
                  <tr>
                    <th className="approvalform-tableCol-width-1" style={{ fontSize: "13px" }}>Received Date</th>
                    <th className="approvalform-tableCol-width-1" style={{ fontSize: "13px" }}>Action Taken Date</th>
                    <th className="approvalform-tableCol-width-2" style={{ fontSize: "13px" }}>TAT (Hours)</th>
                  </tr>
                  <tr>
                    <td className="approvalform-tableCol-width-2" style={{ fontSize: "13px" }}>
                      {getdateconversion(selectedComment.rawReceivedDate)}
                    </td>
                    <td className="approvalform-tableCol-width-2" style={{ fontSize: "13px" }}>
                      {getdateconversion(selectedComment.rawActionDate)}
                    </td>
                    <td className="approvalform-tableCol-width-2" style={{ fontSize: "13px" }}>
                      {selectedComment.individualTatDays}
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>
            <DialogActionsBar>
              <Button className="notifyDailogOkBtn k-button k-button-md k-rounded-md k-button-solid k-button-solid-base" onClick={() => setIsDialogOpen(false)}>
                <span className="k-icon k-font-icon k-i-check cursor allIconsforPrimary-btn" />
                <span>Ok</span>
              </Button>
            </DialogActionsBar>
          </Dialog>
        )}
    </div >
  );
};




