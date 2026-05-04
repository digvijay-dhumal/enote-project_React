import React, { useState, useEffect } from "react";
// import kendo components
import {
  Form,
  Field,
  FormElement,
  FieldWrapper,
} from "@progress/kendo-react-form";
import { Input, TextArea } from "@progress/kendo-react-inputs";
import { DatePicker } from "@progress/kendo-react-dateinputs";
import { Button } from "@progress/kendo-react-buttons";
import { Dialog, DialogActionsBar } from "@progress/kendo-react-dialogs";
import { DropDownList } from "@progress/kendo-react-dropdowns";
import { Hint } from "@progress/kendo-react-labels";
// import kendo icons and images
import { infoCircleIcon } from "@progress/kendo-svg-icons";
import { SvgIcon } from "@progress/kendo-react-common";
import {
  filePdfIcon,
  fileWordIcon,
  fileImageIcon,
  fileTxtIcon,
  fileDataIcon,
  fileIcon,
} from "@progress/kendo-svg-icons";
import view from "../assets/view.png";
import hide from "../assets/hide.png";
// import external components
import { useMsal, useAccount } from "@azure/msal-react";
import { useNavigate, useParams } from "react-router-dom";
import DateObject from "react-date-object";
import CryptoJS from "crypto-js";
import { v4 as uuidv4 } from 'uuid';
// import internal components
import { PageLoader } from "../components/pageLoader";
import { getAccessToken, useTabContext } from "../App";
import { loginRequest, API_COMMON_HEADERS, API_BASE_URL, API_ENDPOINTS, chunkSize, sanitizeInput, charValidation } from "../config";
import Navbar from "../components/navbar";
import { Sidebar } from "../components/sidebar";
import { MenuNavContainer } from "../components/menu";
import Footer from "../components/footer";
import Unauthorized from "../components/Unauthorized";
// CSS Styles
import "../styles/passcode.css";
import "bootstrap/dist/css/bootstrap.min.css";
import "../styles/forms.css";

const CustomDialogTitleBar = () => {
  return (
    <div className="custom-title">
      <SvgIcon icon={infoCircleIcon} /> Alert!
    </div>
  );
};

const CustomConfirmDialogTitleBar = () => {
  return (
    <div className="custom-title cstDailogIbHeader">
      <span className="k-icon k-font-icon k-i-borders-show-hide cursor allIconsforPrimary-btn"></span>{" "}
      Confirmation
    </div>
  );
};

export const ENoteAtrworkflowform = () => {
  const { setPasscodeNavigate } = useTabContext();
  const [isEditMode] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const [crtAtrObj, setcrtAtrObj] = useState({});
  const [enumObj, setEnumObj] = useState();
  const [updateState, setupdateState] = useState(false);
  const [actionTaken, setactionTaken] = useState("");
  const [actionTakenDt, setactionTakenDt] = useState(new Date());
  const [comments, setcomments] = useState("");
  const id = useParams();
  const { accounts, instance } = useMsal();
  const account = useAccount(accounts[0] || {});
  const [dialogOpen, setDialogOpen] = useState(false);
  const [dialogSuccess, setdialogSuccess] = useState(false);
  const [dialogFailure, setdialogFailure] = useState(false);
  const [isAssignee, setIsAssignee] = useState(false);
  const [isRequester, setIsRequester] = useState(false);
  const navigate = useNavigate();
  const [requesterStatusType, setrequesterStatusType] = useState("Accepted");
  const [redrirectTo] = useState("/atrviews/Completed%20ATR");
  const [failureDailogTxt, setFailureDailogTxt] = useState("");
  const [failuremsg] = useState("Something went  wrong please try again!");
  const [actionTakenError, setActionTakenError] = useState("");
  const [reqCommentsError, setReqCommentsError] = useState("");
  const [isPendingWith, setIsPendingWith] = useState("");
  const [actionTakenDateError, setActionTakenDateError] = useState("");
  const [actionTakenBorderColor, setActionTakenBorderColor] = useState("rgba(0, 0, 0, 0.12)");
  const [actionTakenDateBorderColor, setActionTakenDateBorderColor] = useState("rgba(0, 0, 0, 0.12)");
  const [reqCommentsBorderColor, setReqCommentsBorderColor] = useState("rgba(0, 0, 0, 0.12)");
  const [reviewerBorderColor, setReviewerBorderColor] = useState("rgba(0, 0, 0, 0.12)");
  const [confirmDailogObj, setConfirmDailogObj] = useState({
    Confirmtext: "",
    Description: "",
  });

  const handleCloseDlg = async () => {
    setDialogOpen(!dialogOpen);
    setActionInProgress(false);
  };

  const [supportDocWarning, SetSupportDocWarning] = useState({
    "file.txt": {
      filename: "",
      isValid: false,
      warningMsg: "",
      fileExtnsion: "",
    },
  });

  // supportDoc files info
  const [filesInfo, SetFileinfo] = useState([]);
  const [supportingDocError, setSupportingError] = useState("");
  // Passcode validation and input
  const [passcodeVerification, setPasscodeVerification] = useState(false);
  const [isPasscodeVisible, setIsPasscodeVisible] = useState(false);
  const [passcode, setPasscode] = useState("");
  const [actionBtn, setActionBtn] = useState("");
  const [error, setError] = useState("");
  const [validPasscode, setValidPasscode] = useState(false);
  const [validMsg, setValidmsg] = useState(false);
  const [verifyPasscode, setVerifypasscode] = useState("false");
  const [isVieweEligible, setIsVieweEligible] = useState("true");
  const [reviewerList, setReviewerList] = useState([]);
  const [defaultAssignee, setDefaultAssignee] = useState("");
  const [noteID, setNoteID] = useState("");
  const [reviewerSuccessDialog, setReviewerSuccessDialog] = useState(false);
  const [reviewerStatus, setReviewerStatus] = useState("");
  const [atrReviewersDTO, setAtrReviewersDTO] = useState([]);
  const [isReviewer, setIsReviewer] = useState(false);
  const [selectedReviewerEmail, setSelectedReviewerEmail] = useState("");
  const [reviewerEmailDetails, setReviewerEmailDetails] = useState([]);
  const [selectedApproverName, setSelectedApproverName] = useState("");
  const [externalType, setExternalType] = useState(false);
  // file upload
  // const [fileQueue, setFileQueue] = useState([]);
  // const chunkSize = 1048576 * 3;
  const [notePrefix, setNotePrefix] = useState(""); // State to store the file
  const max = 250;
  const maxActionTaken = 1000;
  const [actionInProgress, setActionInProgress] = useState(false);
   const [isSubmitting, setIsSubmitting] = useState(false);

  useEffect(() => {
    setIsLoading(true);
    isUserViewEligible();
    VerifyUserPasscode();

    const fetchApiData = async (id) => {
      const accessToken = await getAccessToken(
        { ...loginRequest, account },
        instance
      );
      //get enum objects
      const dropdowns = await fetch(
        `${API_BASE_URL}${API_ENDPOINTS.GET_DROPDOWNDATA}`,
        {
          method: "GET",
          headers: {
            ...API_COMMON_HEADERS,
            Authorization: `Bearer ${accessToken}`,
          },
        }
      );
      const enumsObj = await dropdowns.json();
      setEnumObj(enumsObj);
      const isExternalAtrType = enumsObj.ATRType?.some(type => type.dValue === "External" && type.id === crtAtrObj.atrType);
      setExternalType(isExternalAtrType);

      await fetch(`${API_BASE_URL}${API_ENDPOINTS.ATR_GetRequest}`, {
        method: "POST",
        body: JSON.stringify({
          AtrId: id,
        }),
        headers: {
          ...API_COMMON_HEADERS,
          Authorization: `Bearer ${accessToken}`,
        },
      })
        .then((res) => {
          return res.json();
        })
        .then((data) => {
          const notePrefix = data.noteNumber.split('/')[0];  // This will give you "MW&AD"
          setNotePrefix(notePrefix);  // Store the value in state, assuming you have a setNotePrefix function

          setupdateState(true);
          setNoteID(data.noteId);
          getReviewerDetails(data.noteId, data.atrId);
          setcrtAtrObj(data);
          setactionTaken(crtAtrObj.actionTaken || "");
          setactionTakenDt(
            data.actionTakenDate
              ? new Date(crtAtrObj.actionTakenDate)
              : new Date()
          );
          setcomments(data.noteRequesterComments || "");

          setIsAssignee(
            (
              // 1. Check if the status is either "Pending Assignee" or "Returned"
              crtAtrObj.atrStatus === enumsObj.ATRStatusEnum.find((x) => x.id === 1).id ||
              crtAtrObj.atrStatus === enumsObj.ATRStatusEnum.find((x) => x.id === 3).id
            ) &&
            // 2. Ensure that data.atrApproversDTO is not an empty array when atrType is not 2

            // 3. Check if the current user is either the assignee or the default assignee
            (data.atrAssigneeEmail === accounts[0].username || data.defaultAssignee === accounts[0].username)
          );

          setIsReviewer(
            (crtAtrObj.atrStatus === enumsObj.ATRStatusEnum.find((x) => x.id === 5).id ||
              crtAtrObj.atrStatus === enumsObj.ATRStatusEnum.find((x) => x.id === 1).id ||
              crtAtrObj.atrStatus === enumsObj.ATRStatusEnum.find((x) => x.id === 3).id) &&
            data.atrAssigneeEmail === accounts[0].username
          );

          setDefaultAssignee(data.defaultAssignee);
          setAtrReviewersDTO(data.atrReviewersDTO);
          const reviewerEmail = data.atrReviewersDTO.map(obj => obj.atrReviewerEmail);
          setReviewerEmailDetails(reviewerEmail);

          setIsRequester(
            (
              // 1. Check if the current user is a reviewer and the status is "Pending Reviewer"
              (data.atrReviewersDTO.some((reviewer) => reviewer.atrReviewerEmail === accounts[0].username) && data.strAtrStatus === "Pending ATR Review") ||
              // 2. Check if the current user is an approver
              (data.atrApproversDTO.some((approver) => approver.approverEmail === accounts[0].username && approver.strStatus === "Pending")) ||
              // 3. Check if the current user is the secretary and the status is "Completed"
              (data.secrataryEmail === accounts[0].username && data.strAtrStatus === "Action Taken")
            )
          );

          //Reviewer name update logic added  - Venkat Oct 19

          setIsPendingWith(
            data.atrStatus === enumsObj.ATRStatusEnum.find((x) => x.id === 4).id
              ? ""
              : data.atrStatus === enumsObj.ATRStatusEnum.find((x) => x.id === 2).id
                ? data.secrataryEmailName
                : data.strAtrStatus === "Pending ATR Review"
                  ? (data.atrApproversDTO[0]?.strApproverEmailName ?? data.atrReviewersDTO[0]?.atrReviewerEmailName)
                  : data.atrAssignerEmailName
          );

          if (data.atrSupportDocumentsDTO.length > 0) {
            let upFileObj = data?.atrSupportDocumentsDTO?.map((x) => ({
              supportingDocumentFileName: x.supportingDocumentFileName,
              supportDocumentPath: x.supportDocumentPath,
              supportDocumentPathLength: x.supportDocumentPathLength,
            }));
            SetFileinfo(upFileObj);
          }
          setIsLoading(false);
        })
        .catch((err) => {
          console.error("Error fetching data:", err);
          setIsLoading(false);
        });
    };
    fetchApiData(id.id); // Set the default noteId
  }, [updateState]);

  const getReviewerDetails = async (noteId, atrID) => {
    try {
      const accessToken = await getAccessToken({ ...loginRequest, account }, instance);

      const reviewerresponse = await fetch(
        `${API_BASE_URL}${API_ENDPOINTS.atrReviewerDropdown}?NoteId=${noteId}&AtrId=${atrID}`, {
        method: "GET",
        headers: { ...API_COMMON_HEADERS, Authorization: `Bearer ${accessToken}` }
      });
      const reviewerList = await reviewerresponse.json();
      if (reviewerList.length === 0) {
        setReviewerList([]);
      } else {
        setReviewerList(reviewerList);
      }

    } catch (error) {
      console.error("Error fetching user details:", error);
    }
  };

  const isUserViewEligible = async () => {
    let isUserViewEligible = true;
    try {
      const accessToken = await getAccessToken(
        { ...loginRequest, account },
        instance
      );
      const response = await fetch(
        `${API_BASE_URL}${API_ENDPOINTS.eNote_ATRViewEligibility}?userPrincipalName=${accounts[0].username}&AtrId=${id.id}`, {
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
          setIsLoading(false);
        }

        setIsVieweEligible(result)
      }
      return isUserViewEligible;
    } catch (error) {
      setIsLoading(false);
      console.error("Error fetching data:", error);
    }
  };

  // Hanlde multiple doc attchment controll

  // const multipleDocUpload = (e, fileCategory) => {
  //   let isValidFile = true;
  //   let warningMsg = "";
  //   let remainingFileSize = 26214400; // 25 MB limit
  //   const inValidFileNames = Object.keys(supportDocWarning).filter(
  //     (fileName) => supportDocWarning[fileName]?.isValid === false
  //   );

  //   if (filesInfo.length > 0) {
  //     const validMultipleFile = filesInfo.filter(obj => {
  //       if (!inValidFileNames.includes(obj.supportingDocumentFileName)) {
  //         return obj;
  //       }
  //     });
  //     validMultipleFile.forEach(obj =>
  //       remainingFileSize -= obj.supportingDocumentPathLength
  //     );
  //   }

  //   const arrayExtension = [
  //     ".pdf", ".doc", ".docx", ".xlsx",
  //     ".PDF", ".DOC", ".DOCX", ".XLSX"
  //   ];

  //   const validNamePattern = /[!@#$%^&*(){}\[\];:,<>\?\/\\]/;
  //   const selectedFiles = Array.from(e.target.files);

  //   const tempFileInfo = [...filesInfo];
  //   const newFileQueue = [];
  //   const newProgressMap = {};

  //   selectedFiles.forEach(file => {
  //     const fileExtension = file.name.split('.').pop();

  //     if (!filesInfo.some(obj => obj.supportingDocumentFileName === file.name)) {
  //       if (!arrayExtension.includes(`.${fileExtension}`)) {
  //         isValidFile = false;
  //         warningMsg = "File type is not allowed";
  //         SetSupportDocWarning({
  //           ...supportDocWarning,
  //           [file.name]: {
  //             filename: file.name,
  //             isValid: false,
  //             warningMsg,
  //             fileExtension
  //           },
  //         });
  //       } else if (validNamePattern.test(file.name)) {
  //         warningMsg = "File name should not contain special characters";
  //         SetSupportDocWarning({
  //           ...supportDocWarning,
  //           [file.name]: {
  //             filename: file.name,
  //             isValid: false,
  //             warningMsg,
  //             fileExtension
  //           },
  //         });
  //       } else if (file.size > remainingFileSize) {
  //         warningMsg = "Cumulative size of all the supporting documents should not exceed 25 MB.";
  //         setSupportingError(warningMsg);
  //       } else {
  //         // Valid file, add to the queue and initialize progress
  //         newFileQueue.push({
  //           file,
  //           category: fileCategory
  //         });
  //         newProgressMap[file.name] = 0; // Initialize progress
  //       }
  //     }
  //   });

  //   // Update file queue and progress map
  //   setFileQueue(prevQueue => [...prevQueue, ...newFileQueue]);

  //   // Create new file info objects
  //   const updatedTempFileInfo = newFileQueue.map(fileObj => ({
  //     noteSupportingDocumentId: 0,
  //     noteId: 0,
  //     supportingDocumentFileName: fileObj.file.name,
  //     createdDate: new Date(),
  //     createdBy: accounts[0].username,
  //     modifiedDate: new Date(),
  //     modifiedBy: accounts[0].username,
  //     supportingDocumentPathLength: fileObj.file.size
  //   }));

  //   SetFileinfo([...tempFileInfo, ...updatedTempFileInfo]);
  // };


  // Function to process and send file chunks
  const sendChunks = async (file, fileCategory, noteID, notePrefix) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = async (event) => {
        const base64Data = event.target.result.split(",")[1];
        if (!base64Data) return reject("Base64 data is undefined.");

        const totalChunks = Math.ceil(base64Data.length / chunkSize);
        for (let i = 0; i < totalChunks; i++) {
          const chunkData = base64Data.slice(i * chunkSize, (i + 1) * chunkSize);
          const response = await sendChunkToServer(chunkData, i, totalChunks, fileCategory, noteID, notePrefix, file.name);

          if (!response.ok) return reject(`Error uploading chunk: ${response.statusText}`);
        }
        resolve("Upload successful");
      };

      reader.onerror = (error) => reject(error);
      reader.readAsDataURL(file);
    });
  };

  const sendChunkToServer = async (NotePdfPathPart1, FileId, NotePdfPathPart2, fileCategory, noteID, departmentAlias, fileName) => {
    if (!NotePdfPathPart1) {
      console.error("Chunk data is undefined.");
      return { status: 400, statusText: "Chunk data is undefined." };
    }
    console.log(crtAtrObj, "crtAtrObj");


    const data = {
      NotePdfPathPart1: String(NotePdfPathPart1),
      FileId: String(FileId),
      NotePdfPathPart2: String(NotePdfPathPart2),
      fileName: fileName,
      Category: "Enote",
      FileCategory: "SupportingDocument",
      DeptAlieasName: departmentAlias,
      NoteId: noteID,
      atrID: id.id,
      AtrCreatedBy: crtAtrObj.createdBy,
    };

    try {
      let accessToken = await getAccessToken({ ...loginRequest, account }, instance);

      const response = await fetch(`${API_BASE_URL}${API_ENDPOINTS.UploadBase64ATR}`, {
        method: "POST",
        headers: { ...API_COMMON_HEADERS, Authorization: `Bearer ${accessToken}` },
        body: JSON.stringify(data),
      });

      if (response.ok) {
        console.log(`Chunk ${FileId + 1}/${NotePdfPathPart2} uploaded successfully.`);
      } else if (response.status === 403) {
        console.error("403 Forbidden: Access is denied.");
      } else {
        console.error(`Error uploading chunk: ${response.statusText}`);
      }
      return response;
    } catch (error) {
      console.error("Error uploading chunk:", error);
      return { status: 500, statusText: error.message };
    }
  };

  const multipleDocUpload = async () => {
    // Added passcode redirect page --> Kavya (23/07) feedback -- 427
    if (verifyPasscode === "true") {
      let waringmsg = "";
      let totalFileSize = 26214400;
      const inValidFileNames = Object.keys(supportDocWarning).filter(
        (fileName) => supportDocWarning[fileName]?.isValid === false
      );

      if (filesInfo.length > 0) {
        const vaildMultiplefile = filesInfo.filter((obj) => {
          if (!inValidFileNames.includes(obj.supportingDocumentFileName)) {
            return obj;
          }
        });
        vaildMultiplefile.map(
          (obj) =>
            (totalFileSize = totalFileSize - obj.supportDocumentPathLength)
        );
      }

      const arrayExtension = [
        ".pdf",
        ".doc",
        ".docx",
        ".xlsx",
        ".PDF",
        ".DOC",
        ".DOCX",
        ".XLSX",
      ];
      const validname = /[!@#$%^&*(){}\[\];:,<>\?\/\\]/;
      const selectedFile = document.getElementById("multiDoc").files;
      const TempfileInfo = filesInfo;

      const promises = [];
      const fileCount = selectedFile.length;

      for (let i = 0; i < fileCount; i++) {
        const fileToLoad = selectedFile[i];
        const fileExtession = fileToLoad.name.split(".");

        if (validname.test(fileToLoad.name)) {
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
        } else if (
          !arrayExtension.includes(
            `.${fileExtession[fileExtession.length - 1]}`
          )
        ) {
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
        } else if (
          fileToLoad.size > totalFileSize &&
          !filesInfo?.some(
            (obj) => obj.supportingDocumentFileName === fileToLoad.name
          )
        ) {
          waringmsg =
            "Cumulative size of all the supporting documents should not be exceeded 25 MB.";
          setSupportingError(waringmsg);
        }

        const fileReader = new FileReader();
        const promise = new Promise((resolve) => {
          fileReader.onload = function (fileLoadedEvent) {
            resolve({
              name: fileToLoad.name,
              base64: fileLoadedEvent.target.result,
              size: fileToLoad.size,
            });
          };
        });

        fileReader.readAsDataURL(fileToLoad);
        promises.push(promise);
      }

      Promise.all(promises)
        .then((fileDataArray) => {
          const updatedTempFileInfo = fileDataArray.reduce(
            (acc, fileData) => {
              const ObjExist = acc.map((obj) => obj.supportingDocumentFileName);
              if (!ObjExist.includes(fileData.name)) {
                // Split the base64 string into 10 parts and name them accordingly
                const base64Parts = splitBase64IntoNamedParts(
                  fileData.base64.split(",")[1],
                  10
                );
                acc.push({
                  ...base64Parts,
                  supportingDocumentFileName: fileData.name,
                  supportDocumentPathLength: fileData.size,
                });
              }
              return acc;
            },
            [...TempfileInfo]
          );
          SetFileinfo(updatedTempFileInfo);
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

  // Handle split Base64 into named parts
  //  const [uploadedFiles, setUploadedFiles] = useState([]); // Store files for processing
  // const multipleDocUpload = async () => {
  //   // Added passcode redirect page --> Kavya (23/07) feedback -- 427
  //   if (verifyPasscode !== "true") {
  //     setValidPasscode(true);
  //     setValidmsg("Passcode is not set. Please create passcode to proceed further.");
  //     return;
  //   }

  //   let totalFileSize = 26214400; // 25MB limit
  //   let warningMsg = "";

  //   const inValidFileNames = Object.keys(supportDocWarning).filter(
  //     (fileName) => supportDocWarning[fileName]?.isValid === false
  //   );

  //   if (filesInfo.length > 0) {
  //     const validExistingFiles = filesInfo.filter(
  //       (obj) => !inValidFileNames.includes(obj.supportingDocumentFileName)
  //     );

  //     validExistingFiles.forEach(
  //       (obj) => (totalFileSize -= obj.supportDocumentPathLength)
  //     );
  //   }

  //   const allowedExtensions = ["pdf", "doc", "docx", "xlsx"];
  //   const specialCharPattern = /[!@#$%^&*(){}\[\];:,<>\?\/\\]/;
  //   const selectedFiles = document.getElementById("multiDoc").files;
  //   const newFiles = [];

  //   if (selectedFiles.length === 0) {
  //     setSupportingError("Please select at least one file.");
  //     return;
  //   }

  //   let updatedFileInfo = [...filesInfo];

  //   for (let i = 0; i < selectedFiles.length; i++) {
  //     const file = selectedFiles[i];
  //     const fileName = file.name;
  //     const fileExtension = fileName.split(".").pop().toLowerCase();
  //     const fileSize = file.size;

  //     // Validate file name special characters
  //     if (specialCharPattern.test(fileName)) {
  //       warningMsg = `File name should not contain special characters: ${fileName}`;
  //       updateFileWarnings(fileName, warningMsg,fileExtension);  
  //     }

  //     // Validate file extension
  //     if (!allowedExtensions.includes(fileExtension)) {
  //       warningMsg = `File type is not allowed: ${fileName}`;
  //       updateFileWarnings(fileName, warningMsg,fileExtension);
  //     }

  //     // Validate cumulative file size
  //     if (fileSize > totalFileSize) {
  //       warningMsg = "Cumulative size of all supporting documents should not exceed 25 MB.";
  //       // setSupportingError(warningMsg);
  //       updateFileWarnings(fileName, warningMsg,fileExtension);
  //       continue;
  //     }

  //     // Prevent duplicate files
  //     if (filesInfo.some((obj) => obj.supportingDocumentFileName === fileName)) {
  //       continue;
  //     }

  //     totalFileSize -= fileSize;

  //     // Add valid files
  //     newFiles.push(file);
  //     updatedFileInfo.push({
  //       supportingDocumentFileName: fileName,
  //       createdDate: new Date(),
  //       createdBy: accounts[0].username,
  //       modifiedDate: new Date(),
  //       modifiedBy: accounts[0].username,
  //       supportingDocumentPathLength: fileSize,
  //     });
  //   }


  //   SetFileinfo(updatedFileInfo);
  //   setUploadedFiles(newFiles); // Store valid files for later processing
  // };

  // const updateFileWarnings=(fileName, warningMsg,fileExtension)=>{
  //   SetSupportDocWarning({
  //     ...supportDocWarning,
  //     [fileName]: {
  //       filename: fileName,
  //       isValid: true,
  //       warningMsg:warningMsg,
  //       fileExtension,
  //     },
  //   })
  // }

  const splitBase64IntoNamedParts = (base64String, parts) => {
    const partSize = Math.ceil(base64String.length / parts);
    const base64Parts = {};
    for (let i = 0; i < parts; i++) {
      const partName = `supportDocumentPathPart${i + 1}`;
      base64Parts[partName] = base64String.slice(
        i * partSize,
        (i + 1) * partSize
      );
    }
    return base64Parts;
  };

  /* Validates Form Data */
  const validateForm = () => {
    const errors = {};
    let totalFileSize = 0;
    const inValidFileNames = Object.keys(supportDocWarning).filter(
      (fileName) => supportDocWarning[fileName]?.isValid === false
    );
    // filter valid files only if()
    if (filesInfo.length > 0) {
      const vaildMultiplefile = filesInfo.filter((obj) => {
        if (!inValidFileNames.includes(obj.supportingDocumentFileName)) {
          return obj;
        }
      });
      vaildMultiplefile.map(
        (obj) => (totalFileSize = totalFileSize + obj.supportDocumentPathLength)
      );
    }
    if (Object.keys(supportDocWarning).length > 1 || totalFileSize > 26214400) {
      if (Object.keys(supportDocWarning).length > 1) {
        errors["Supporting Documents"] = "Please select vaild files";
      }
      if (totalFileSize > 26214400) {
        errors["Supporting DocumentsMaxSize"] =
          "Cumulative size of all the supporting documents should not be exceeded 25 MB.";
      }
    }

    if (!actionTaken.trim()) {
      setActionTakenError("Action Taken field should not be empty.");
      errors["ActionTaken"] = "NoteTo field should not be empty.";
      setActionTakenBorderColor("#f31700");
    } else {
      setActionTakenError("");
      setActionTakenBorderColor("");
    }
    if (!actionTakenDt) {
      setActionTakenDateError("Action Taken Date field should not be empty.");
      errors["ActionTakenDate"] =
        "Action Taken Date field should not be empty.";
      setActionTakenDateBorderColor("#f31700");
    } else {
      setActionTakenDateError("");
      setActionTakenDateBorderColor("");
    }
    if (!selectedApproverName && !crtAtrObj.atrReviewersDTO.length > 0 && reviewerList.length > 0) {
      setReviewerBorderColor("#f31700");
      errors["Reviewer"] = "Reviewer field should not be empty.";
    } else {
      setReviewerBorderColor("");

    }
    return errors;
  };

  // checking file type and return file icon
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

  const handleReviewerChange = (e) => {
    const value = e.target.value;
    setReviewerBorderColor("");
    const selectedReviewer = reviewerList.find((x) => x.approverEmailName === value);
    setSelectedReviewerEmail(selectedReviewer.approverEmail);
    setSelectedApproverName(selectedReviewer.approverEmailName)

    setcrtAtrObj({
      ...crtAtrObj,
      approverEmail: selectedReviewer.approverEmail,
      approverName: selectedReviewer.approverEmailName,
    });
  };

  // Handle on change action taken
  const handleOnChangeActionTaken = async (event) => {
    
    setActionInProgress(true);
    try{
    // Added passcode redirect page --> Kavya (23/07) feedback -- 427
    let value = event.target.value;
    // Sanitize the input value
    // value = sanitizeInput(value);

    // const allowedCharsRegex = charValidation;

    if (verifyPasscode === "true") {
      // if (allowedCharsRegex.test(value)) {
      setactionTaken(value);
      setActionTakenBorderColor("");
      setActionTakenError("");
      console.log(actionTaken);
      // } else {
      //   setActionTakenBorderColor("#f31700");
      //   setActionTakenError("special characters are not allowed.");
      // }
    } else {
      setValidPasscode(true);
      // Changed the content --> Kavya (25-07)
      setValidmsg(
        "Passcode is not set. Please create passcode to proceed further."
      );
    }
  } finally {
    setActionInProgress(false);
  }
  };

  // Handle on change action taken Date
  const handleOnChangeActionTakenDt = async (event) => {
    // Added passcode redirect page --> feedback 427 Kavya (23/07)
    if (verifyPasscode === "true") {
      setactionTakenDt(event.target.value);
      setActionTakenDateError("");
      setActionTakenDateBorderColor("");
      console.log(actionTaken);
    } else {
      setValidPasscode(true);
      // Changed the content --> Kavya (25-07)
      setValidmsg(
        "Passcode is not set. Please create passcode to proceed further."
      );
      console.log(actionTaken);
    }
  };

  // Handle close dialog
  const handleClsDlgSuccess = () => {
    setdialogSuccess(!dialogSuccess);
    if (isAssignee) {
      navigate("/atrviews/Pending%20ATR");
    }
    if (isRequester) {
      navigate("/atrviews/Pending Approval");
    }
  };

  // It's help redirect the page based on condtion
  const handleDlgBtnSuccess = () => {
    setdialogSuccess(!dialogSuccess);
    if (isAssignee) {
      navigate("/atrviews/Pending%20ATR");
    }
    if (isRequester) {
      navigate("/atrviews/Pending Approval");
    }
  };

  // Handle close dialog for failure
  const handleClsDlgFailure = () => {
    setdialogFailure(!dialogFailure);
  };

  // Handle hide dailog
  const handleDlgBtnFailure = () => {
    setdialogFailure(!dialogFailure);
    setError("");
  };

  // Handle ATR comments change
  const handleatrComments = async (event) => {
    // Added passcode redirect page --> feedback 427 Kavya (23/07)
    const value = event.target.value;

    if (verifyPasscode === "true") {
      // if (allowedCharsRegex.test(value)) {
      setcomments(value);
      setReqCommentsError("");
      setReqCommentsBorderColor(""); // Optional: reset border color on valid input
        if (value.trim() !== "") {
        setActionInProgress(false);
      }
      // } else {
      //   setReqCommentsError("Special characters are not allowed.");
      //   setReqCommentsBorderColor("#f31700");
      // }
    } else {
      setValidPasscode(true);
      // Changed the content --> Kavya (25-07)
      setValidmsg(
        "Passcode is not set. Please create passcode to proceed further."
      );
    }
  };


  // Handle Exit
  const handleExit = () => {
    navigate(redrirectTo);
  };
  // supportDoc link
  //Convert base64 to URL - supportDoc
  const getSupportDocHyperlink = async (filePath, fileName) => {
    const accessToken = await getAccessToken(
      { ...loginRequest, account },
      instance
    );
    setIsLoading(true);
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
    setIsLoading(false);
  };

  // getfileType
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

  // remove multiple attchemnts
  const onRemoveMultiAttachment = (id) => {
    const filename = filesInfo.find(
      (obj, index) => index === id
    ).supportingDocumentFileName;
    delete supportDocWarning[filename];
    SetSupportDocWarning(supportDocWarning);
    let totalFileSize = 0;
    const inValidFileNames = Object.keys(supportDocWarning).filter(
      (fileName) => supportDocWarning[fileName]?.isValid === false
    );
    // filter valid files only if()
    if (filesInfo.length > 0) {
      const vaildMultiplefile = filesInfo.filter((obj) => {
        if (!inValidFileNames.includes(obj.supportingDocumentFileName)) {
          return obj;
        }
      });
      vaildMultiplefile.map((obj, index) => {
        if (index !== id) {
          totalFileSize = totalFileSize + obj.supportDocumentPathLength;
        }
      });

      if (totalFileSize <= 26214400) {
        setSupportingError("");
      }
    }
    SetFileinfo(filesInfo.filter((obj, ind) => ind !== id));
  };

  // onchange passcode
  const handlePasscodeChange = (e) => {
    const { value } = e.target;
    // Only allow numeric input
    const alphanumericRegex = /^[a-zA-Z0-9]*$/;

    if (alphanumericRegex.test(value)) {
      setPasscode(value);
    }
  };

  // Handle Redirect to Passcode page
  const handleRedirectToPasscode = () => {
    setPasscodeNavigate("View");
    navigate("/enotepasscode"); // Adjust the path to your target route
  };

  // Verifing the passcode
  const passcodeVerificationFunction = async () => {
    try {
      const accessToken = await getAccessToken(
        { ...loginRequest, account },
        instance
      );
      // Replace with your own secret key
      const secretKey = "SHA256";

      // Compute SHA256 hash of the passcode with the secret key
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
        setPasscodeVerification(false);
        setConfirmDailogObj({
          Confirmtext: "Are you sure you want to Complete this ATR?",
          Description:
            "Please check the details filled along with attachment and click on confirm button to complete the ATR.",
        });
        setDialogOpen(!dialogOpen);
        if (actionBtn === "Accept") {
          setConfirmDailogObj({
            Confirmtext: "Are you sure you want to Accept this action taken?",
            Description:
              "Please check the details filled along with attachment and click on confirm button to accept the action taken.",
          });
          setrequesterStatusType("Accepted");
          setDialogOpen(!dialogOpen);
        }
        if (actionBtn === "Return") {
          setConfirmDailogObj({
            Confirmtext:
              "Are you sure you want to return this ATR to assignee?",
            Description:
              "Please click on Confirm button to return this ATR to assignee.",
          });

          setReqCommentsError("");
          setReqCommentsBorderColor("");
          setrequesterStatusType("Returned");
          setDialogOpen(!dialogOpen);
        }
      } else {
        setError("Incorrect passcode entered please try again.");
      }
    } catch (err) {
      console.error(err);
      // Handle error state or throw it further
      throw err;
    }
  };

  // Verify User havve Passcode or not
  const VerifyUserPasscode = async () => {
    let userPasscodeExist = "false";
    try {
      const accessToken = await getAccessToken(
        { ...loginRequest, account },
        instance
      );
      // Call eNote_VerifyUserPasscode API first
      const userPasscodeResponse = await fetch(
        `${API_BASE_URL}${API_ENDPOINTS.eNote_VerifyUserPasscode(
          accounts[0].username
        )}`,
        {
          method: "POST",
          // body: JSON.stringify({ UserMail: accounts[0].username}),
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
      // Handle error state or throw it further
      throw err;
    }
    return userPasscodeExist;
  };

  //  Handle passcode dialog close
  const handlepasscodeClose = () => {
    setPasscodeVerification(false);
    setActionInProgress(false);
    setError(false);
  };

  // Handle action complete

  const handleBtnActionComplete = async () => {
    setActionInProgress(false);
    try {
       
    let isFilesValid = true;
    filesInfo.map((x) => {
      isFilesValid = !(x.supportingDocumentFileName in supportDocWarning);
    });

    let totalFileSize = 0;
    const inValidFileNames = Object.keys(supportDocWarning).filter(
      (fileName) => supportDocWarning[fileName]?.isValid === false
    );
    if (filesInfo.length > 0) {
      const validMultipleFiles = filesInfo.filter((obj) => {
        if (!inValidFileNames.includes(obj.supportingDocumentFileName)) {
          return obj;
        }
      });
      validMultipleFiles.map(
        (obj) => (totalFileSize = totalFileSize + obj.supportDocumentPathLength)
      );
    }

    const validateFields = validateForm();

    if (isFilesValid && Object.keys(validateFields).length === 0) {
      // if (verifyPasscode === "true") {
      //   setPasscodeVerification(true);
      //   setPasscode("");
      // } else {
      //   setValidPasscode(true);
      //   setValidmsg(
      //     "Passcode is not set. Please create passcode to proceed further."
      //   );
      // }
      // --> Commented passcode verification logic --> Kavya (04-11)
      setConfirmDailogObj({
        Confirmtext: "Are you sure you want to Complete this ATR?",
        Description:
          "Please check the details filled along with attachment and click on confirm button to complete the ATR.",
      });
      setDialogOpen(!dialogOpen);
    } else {
      setdialogFailure(true);
      if (totalFileSize > 26214400) {
        setFailureDailogTxt(
          "Cumulative size of all the supporting documents should not be exceeded 25 MB."
        );
      } else {
        setFailureDailogTxt(
          "Please fill the mandatory fields, check and remove if there are any invalid files attached."
        );
      }
    }
  }finally{
    setActionInProgress(false);
  }
  };

  // handle Accept
  const handleBtnAccept = async () => {
    // Added passcode redirect page --> feedback 427 Kavya (23/07)
    //  ---> Commented passcode verification logic --> Kavya (04-11)
    // if (verifyPasscode === "true") {
    //   setActionBtn("Accept");
    //   setPasscodeVerification(true);
    //   setPasscode("");
    // } else {
    //   setActionBtn("Accept");
    //   setValidPasscode(true);
    //   // Changed the content --> Kavya (25-07)
    //   setValidmsg(
    //     "Passcode is not set. Please create passcode to proceed further."
    //   );
    // }

    setConfirmDailogObj({
      Confirmtext: "Are you sure you want to Accept this action taken?",
      Description:
        "Please check the details filled along with attachment and click on confirm button to accept the action taken.",
    });
    setrequesterStatusType("Accepted");
    setDialogOpen(!dialogOpen);
  };

  // handle return btn

  const handleBtnReturn = async () => {
    setActionInProgress(true);
    if (crtAtrObj.strAtrStatus === "Pending ATR Review") {
      // Check if actionTaken is provided when status is "Pending ATR Review"
      if (!actionTaken.trim()) {
        setdialogFailure(true);
        setFailureDailogTxt("Please fill the action taken to return the ATR.");
        setActionTakenError("Action taken field cannot be empty.");
        setActionTakenBorderColor("#f31700");
      } else {
        // Proceed with confirmation dialog if actionTaken is provided
        setConfirmDailogObj({
          Confirmtext: "Are you sure you want to return this ATR to assignee?",
          Description:
            "Please click on Confirm button to return this ATR to assignee.",
        });
        setReqCommentsError("");
        setReqCommentsBorderColor("");
        setrequesterStatusType("Returned");
        setDialogOpen(!dialogOpen);
      }
    } else {
      // If ATR status is not "Pending ATR Review", then check for comments
      if (!comments.trim()) {
        setdialogFailure(true);
        setFailureDailogTxt("Please fill the comments to return the ATR.");
        setReqCommentsError("Comments field cannot be empty.");
        setReqCommentsBorderColor("#f31700");
      } else {
        setConfirmDailogObj({
          Confirmtext: "Are you sure you want to return this ATR to assignee?",
          Description:
            "Please click on Confirm button to return this ATR to assignee.",
        });
        setReqCommentsError("");
        setReqCommentsBorderColor("");
        setrequesterStatusType("Returned");
        setDialogOpen(!dialogOpen);
      }
    }
  };

  // const uploadFiles = (files, enoteID, departmentAlias) => {
  //   const uploadPromises = files.map(({ file, category }) => {
  //     const _fileID = uuidv4() + "." + file.name.split('.').pop(); // Generate unique ID
  //     return uploadChunk(file, _fileID, category, enoteID, departmentAlias); // Pass departmentAlias
  //   });
  //   return Promise.all(uploadPromises); // Wait for all uploads to finish
  // };

  // const uploadChunk = async (file, _fileID, fileCategory) => {
  //   const chunkCount = Math.ceil(file.size / chunkSize);
  //   let counter = 1;
  //   let start = 0;
  //   let end = chunkSize;

  //   while (counter <= chunkCount) {
  //     const chunk = file.slice(start, end);
  //     try {
  //       const accessToken = await getAccessToken({ ...loginRequest, account }, instance);
  //       const response = await fetch(`${API_BASE_URL}${API_ENDPOINTS.eNote_uploadChunk}?id=${counter}&fileId=${_fileID}`, {
  //         method: 'POST',
  //         headers: {
  //           "Content-Type": "application/octet-stream", Authorization: `Bearer ${accessToken}`,
  //         },
  //         body: chunk
  //       });

  //       const data = await response.json();
  //       if (data.isSuccess) {
  //         start = end;
  //         end = start + chunkSize;

  //         counter++;
  //       } else {
  //         console.error('Error uploading chunk:', data.errorMessage);
  //         break;
  //       }
  //     } catch (error) {
  //       console.error('Error during chunk upload:', error);
  //       break;
  //     }
  //   }

  //   // Call uploadCompleted API once all chunks are uploaded
  //   await uploadCompleted(file, _fileID, fileCategory);
  // };
  // const uploadChunk = async (
  //     file,
  //     _fileID,
  //     fileCategory,
  //     enoteID,
  //     departmentAlias,
  //     files,
  //     chunkEnd = chunkSize // Default end
  //   ) => {
  //     const chunkCount = Math.ceil(file.size / chunkSize);
  //     let counter = 1;
  //     let start = 0; // Use the provided start value
  //     let end = chunkEnd; // Use the provided end value

  //     const encryptContent = (content) => {
  //       const key = 0x5A; // Example key
  //       return content.map(byte => byte ^ key);
  //     };

  //     while (counter <= chunkCount) {

  //       // Extract the chunk
  //       const chunk = file.slice(start, end);
  //       const arrayBuffer = await chunk.arrayBuffer(); // Convert Blob to ArrayBuffer
  //       const uint8Array = new Uint8Array(arrayBuffer); // Convert ArrayBuffer to Uint8Array

  //       // Encrypt the chunk
  //       const encryptedChunk = new Uint8Array(encryptContent(uint8Array));

  //       try {
  //         const accessToken = await getAccessToken({ ...loginRequest, account }, instance);
  //         const response = await fetch(
  //           `${API_BASE_URL}${API_ENDPOINTS.eNote_uploadChunk}?id=${counter}&fileId=${_fileID}`,
  //           {
  //             method: 'POST',
  //             headers: {
  //               "Content-Type": "application/octet-stream",
  //               Authorization: `Bearer ${accessToken}`,
  //             },
  //             body: encryptedChunk // Send the encrypted chunk
  //           }
  //         );

  //         const data = await response.json();
  //         if (data.isSuccess) {
  //           start = end;
  //           end = start + chunkSize;
  //           counter++;
  //         } else {
  //           console.error('Error uploading chunk:', data.errorMessage);
  //           break;
  //         }
  //       } catch (error) {
  //         console.error('Error during chunk upload:', error);
  //         break;
  //       }
  //     }
  //     // Check if the file still exists before calling uploadCompleted
  //     await uploadCompleted(file, _fileID, fileCategory);
  //   }

  // const uploadCompleted = async (file, _fileID, fileCategory) => {
  //   try {
  //     const params = {
  //       fileName: file.name,
  //       fileId: _fileID,
  //       fileCategory: fileCategory,
  //       category: 'Enote',
  //       deptAlieasName: notePrefix,
  //       noteID: crtAtrObj.noteId,
  //       AtrID: crtAtrObj.atrId,
  //       AtrCreatedBy: crtAtrObj.createdBy,
  //     }
  //     const accessToken = await getAccessToken({ ...loginRequest, account }, instance);
  //     const response = await fetch(
  //       `${API_BASE_URL}${API_ENDPOINTS.eNote_ATRUploadComplete}`,
  //       {
  //         method: 'POST',
  //         headers: {
  //           "Content-Type": "application/json", Authorization: `Bearer ${accessToken}`,
  //         },
  //         body: JSON.stringify(params),
  //       }
  //     );

  //     const data = await response.json();
  //     if (data.isSuccess) {
  //       console.log('Upload completed for file', file.name);
  //     } else {
  //       console.error('Error completing upload:', data.errorMessage);
  //     }
  //   } catch (error) {
  //     console.error('Error during upload completion:', error);
  //   }
  // };

  //Handle - Action Complete update by Assignee
  // const handleDlgActionComplete = async () => {
  //   setDialogOpen(!dialogOpen);

  //   let actionTkDate = new DateObject(new Date(actionTakenDt)).format("DD/MM/YYYY");

  //   const atrSupportDocumentsDTO = filesInfo?.map((item) => {
  //     const { supportDocumentPathLength, ...rest } = item;
  //     return rest;
  //   });

  //   const atrReviewersDTO = selectedApproverName || crtAtrObj.atrReviewersDTO?.length > 0
  //     ? [
  //       {
  //         ATRCreatorEmail: crtAtrObj.approverEmail,
  //         AtrReviewerEmail: selectedReviewerEmail || crtAtrObj.atrReviewersDTO[0].atrReviewerEmail,
  //         AtrId: crtAtrObj.atrId,
  //         NoteId: crtAtrObj.noteId,
  //         CreatedBy: crtAtrObj.createdBy,
  //       },
  //     ]
  //     : [];

  //   let params = {
  //     atrId: crtAtrObj.atrId,
  //     noteId: crtAtrObj.noteId,
  //     createdBy: accounts[0].username,
  //     atrSupportDocumentsDTO,
  //     ATRReviewersDTO: atrReviewersDTO,
  //     actionTaken: actionTaken,
  //     strActionTakenDate: actionTkDate,
  //     AtrType: crtAtrObj.atrType,
  //     AtrApproverEmail: crtAtrObj.approverEmail,
  //   };

  //   if (atrReviewersDTO.length > 0) {
  //     params.ReviewerComment = crtAtrObj.reviewerComment;
  //     params.atrStatus = 5;
  //   } else {
  //     params.atrStatus = 2;
  //   }

  //   setIsLoading(true);

  //   try {
  //     const accessToken = await getAccessToken({ ...loginRequest, account }, instance);

  //     const response = await fetch(`${API_BASE_URL}${API_ENDPOINTS.ATR_AssigneeAction}`, {
  //       method: "POST",
  //       body: JSON.stringify(params),
  //       headers: {
  //         ...API_COMMON_HEADERS,
  //         Authorization: `Bearer ${accessToken}`,
  //       },
  //     });

  //     const data = await response.json();

  //     if (data.statusMessage === "Success") {
  //       let uploadErrors = [];

  //       // Upload all files and collect errors
  //       for (const file of uploadedFiles) {
  //         try {
  //           await sendChunks(file, "Enote", crtAtrObj.noteId, notePrefix);
  //         } catch (error) {
  //           uploadErrors.push(error);
  //         }
  //       }

  //       setIsLoading(false);

  //       if (uploadErrors.length === 0) {
  //         setdialogSuccess(true);
  //       } else {
  //         setdialogFailure(true);
  //         setFailureDailogTxt("Some files failed to upload. Please try again.");
  //       }
  //     } else {
  //       setIsLoading(false);
  //       setdialogFailure(true);
  //       setFailureDailogTxt(failuremsg);
  //     }
  //   } catch (err) {
  //     console.log(err);
  //     setIsLoading(false);
  //     setdialogFailure(true);
  //     setFailureDailogTxt(failuremsg);
  //   }
  // };

  const handleDlgActionComplete = async () => {
    setActionInProgress(true);
    try{
    setDialogOpen(!dialogOpen);

    // Format the action taken date
    let actionTkDate = new DateObject(new Date(actionTakenDt)).format("DD/MM/YYYY");

    //  document DTOs
    const atrSupportDocumentsDTO = filesInfo?.map((item) => {
      const { supportDocumentPathLength, ...rest } = item;
      return rest; // Keep all other fields except supportDocumentPathLength
    });

    //  ATRReviewersDTO
    const atrReviewersDTO = selectedApproverName || crtAtrObj.atrReviewersDTO?.length > 0
      ? [
        {
          ATRCreatorEmail: crtAtrObj.approverEmail,
          AtrReviewerEmail: selectedReviewerEmail || crtAtrObj.atrReviewersDTO[0].atrReviewerEmail,// Selected from dropdown
          AtrId: crtAtrObj.atrId,
          NoteId: crtAtrObj.noteId,
          CreatedBy: crtAtrObj.createdBy,
        },
      ]
      : []; // Empty array if no reviewer is selected

    // params for the API request
    let params = {
      atrId: crtAtrObj.atrId,
      noteId: crtAtrObj.noteId,
      createdBy: accounts[0].username,
      atrSupportDocumentsDTO: atrSupportDocumentsDTO,
      ATRReviewersDTO: atrReviewersDTO,
      actionTaken: actionTaken,
      strActionTakenDate: actionTkDate,
      AtrType: crtAtrObj.atrType,
      AtrApproverEmail: crtAtrObj.approverEmail,
    };

    // Set status based on whether ATRReviewersDTO has a reviewer
    if (atrReviewersDTO.length > 0) {
      params.ReviewerComment = crtAtrObj.reviewerComment;
      params.atrStatus = 5; // Status for Reviewer if a reviewer is present
    } else {
      params.atrStatus = 2; // Status for Assignee if no reviewer is selected
    }

    // Set loading state
    setIsLoading(true);

    try {
      const accessToken = await getAccessToken({ ...loginRequest, account }, instance);

      // Make the API request
      const response = await fetch(`${API_BASE_URL}${API_ENDPOINTS.ATR_AssigneeAction}`, {
        method: "POST",
        body: JSON.stringify(params),
        headers: {
          ...API_COMMON_HEADERS,
          Authorization: `Bearer ${accessToken}`,
        },
      });
      const data = await response.json();

      // Handle the response
      setIsLoading(false);
      if (data.statusMessage === "Success") {
        setdialogSuccess(true);
      } else {
        setdialogFailure(true);
        setFailureDailogTxt(failuremsg);
      }
    } catch (err) {
      console.log(err);
      setdialogFailure(true);
      setFailureDailogTxt(failuremsg);
    }
  } finally {
    setActionInProgress(false);
  }
  };


  // to remove empty values from object
  const removeEmptyValues = (obj) => {
    return Object.fromEntries(
      Object.entries(obj).filter(
        ([_, value]) => value !== null && value !== '' && value !== 0
      )
    );
  };

  //handle - Accept and Return update by Requester
  const handleDlgRequester = async () => {
    setActionInProgress(true);
    try {
    setDialogOpen(!dialogOpen);
    setActionInProgress(true);
    let params = removeEmptyValues({
      atrId: crtAtrObj.atrId,
      actionerComment: comments,
      createdBy: accounts[0].username,
      RequesterStatus: requesterStatusType, //Accepted/Returned
      actionTaken: actionTaken,
    });
    setIsLoading(!isLoading);

    const accessToken = await getAccessToken(
      { ...loginRequest, account },
      instance
    );

    await fetch(`${API_BASE_URL}${API_ENDPOINTS.ATR_RequesterAction}`, {
      method: "Post",
      body: JSON.stringify(params),
      headers: {
        ...API_COMMON_HEADERS,
        Authorization: `Bearer ${accessToken}`,
      },
    })
      .then((res) => {
        return res.json();
      })
      .then((data) => {
        setIsLoading(false);
        if (data.statusMessage === "Success") {
          setdialogSuccess(true);
        } else {
          setdialogFailure(true);
          setFailureDailogTxt(failuremsg);
        }
      })
      .catch((err) => {
        console.log(err);
        setdialogFailure(true);
        setFailureDailogTxt(failuremsg);
      });
    } finally {
      setActionInProgress(false);
    }
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

  return (
    <>
      <Navbar header="IB Smart Office - eNote" />
      <Sidebar />
      {/* Added top navigation for menu --> 23-09 */}
      <MenuNavContainer isMenuPage />
      {isVieweEligible === "false" ? <> <Unauthorized />  <div className="pgFooterContainer">
        <Footer />
      </div> </> : <>
        <div className="FormMaincontainer cstATRForm">
          <div className="container largeScreen">
            <div className="SectionHeads eNoteAtrFormHdr row mobileSectionHeads">
              {/* ATR Workflow Form */}
              <div className="cstformHdrLeftContainer mobiletitle">
                {crtAtrObj.strAtrStatus === ""
                  ? ""
                  : `Pending With: ${isPendingWith}`}
              </div>
              <div className="cstformHdrMiddleContainer mobiletitle">
                ATR Workflow Form
              </div>
              <div className="cstformHdrRightContainer mobiletitle">
                {`Status: ${crtAtrObj.strAtrStatus}`}
              </div>
            </div>
            <Form
              render={(formRenderProps) => (
                <FormElement>
                  <fieldset className={"k-form-fieldset"}>
                    <div className="SectionRow row">
                      <div className="col-md-6">
                        <FieldWrapper>
                          <div className="k-form-field-wrap">
                            <label htmlFor="noteTo">Note To:</label>
                            <Field
                              component={Input}
                              id="noteTo"
                              name="Noteto"
                              key={crtAtrObj.designation}
                              defaultValue={crtAtrObj.designation}
                              readOnly={!isEditMode}
                            />
                          </div>
                        </FieldWrapper>
                      </div>
                      <div className="col-md-6">
                        <FieldWrapper>
                          <div className="k-form-field-wrap">
                            <label htmlFor="status">Status:</label>
                            <Field
                              component={Input}
                              id="status"
                              name="Status"
                              key={crtAtrObj.strAtrStatus}
                              defaultValue={crtAtrObj.strAtrStatus}
                              readOnly={!isEditMode}
                            />
                          </div>
                        </FieldWrapper>
                      </div>
                      <div className="col-md-6">
                        <FieldWrapper>
                          <div className="k-form-field-wrap">
                            <label htmlFor="creationDate">Creation Date:</label>
                            <Field
                              component={Input}
                              id="creationDate"
                              name="CreationDate"
                              key={crtAtrObj.createdDate}
                              defaultValue={new DateObject(
                                new Date(crtAtrObj.createdDate)
                              ).format("DD-MM-YYYY")}
                              readOnly={!isEditMode}
                            />
                          </div>
                        </FieldWrapper>
                      </div>
                      <div className="col-md-6">
                        <FieldWrapper>
                          <div className="k-form-field-wrap">
                            <label htmlFor="atrNoteid">ATR Note ID:</label>
                            <Field
                              component={Input}
                              id="atrNoteid"
                              name="ATRNoteId"
                              key={crtAtrObj.noteNumber}
                              defaultValue={crtAtrObj.noteNumber}
                              readOnly={!isEditMode}
                            />
                          </div>
                        </FieldWrapper>
                      </div>
                      <div className="col-md-6">
                        <FieldWrapper>
                          <div className="k-form-field-wrap">
                            <label htmlFor="department">Department:</label>
                            <Field
                              component={Input}
                              id="department"
                              name="Department"
                              key={crtAtrObj.departmentName}
                              defaultValue={crtAtrObj.departmentName}
                              readOnly={!isEditMode}
                            />
                          </div>
                        </FieldWrapper>
                      </div>
                      <div className="col-md-6">
                        <FieldWrapper>
                          <div className="k-form-field-wrap">
                            <label htmlFor="subject">Subject:</label>
                            <Field
                              component={Input}
                              id="subject"
                              name="Subject"
                              key={crtAtrObj.subject}
                              defaultValue={crtAtrObj.subject}
                              readOnly={!isEditMode}
                            />
                          </div>
                        </FieldWrapper>
                      </div>
                      <div className="col-md-6">
                        <FieldWrapper>
                          <div className="k-form-field-wrap">
                            <label htmlFor="assingedBy">Assigned By:</label>
                            <Field
                              component={Input}
                              id="assingedBy"
                              name="AssignedBy"
                              key={crtAtrObj.approverEmail}
                              defaultValue={crtAtrObj.approverEmailName}
                              readOnly={!isEditMode}
                            />
                          </div>
                        </FieldWrapper>
                      </div>
                      <div className="col-md-6">
                        <FieldWrapper>
                          <div className="k-form-field-wrap">
                            <label htmlFor="remarks">Remarks:</label>
                            {/* <Field
                              component={Input}
                              id="remarks"
                              name="Remarks"
                              key={crtAtrObj.remarks}
                              defaultValue={crtAtrObj.remarks}
                              readOnly={!isEditMode}
                            /> */}
                            <TextArea
                              id="remarks"
                              name="Remarks"
                              value={
                                Array.isArray(crtAtrObj.remarksList)
                                  ? crtAtrObj.remarksList.join("\n")
                                  : ""
                              }
                              readOnly={!isEditMode}
                              style={{
                                resize: "vertical", // Allow vertical resizing
                                minHeight: "40px"
                              }}
                              rows={3} // Set a reasonable default
                            />
                          </div>
                        </FieldWrapper>
                      </div>

                      {!isRequester && (
                        <div className="col-md-6">
                          <FieldWrapper>
                            <div className="k-form-field-wrap">
                              <label htmlFor="comments">Comments:</label>
                              <TextArea
                                component={Input}
                                maxLength={max}
                                id="comments"
                                name="Comments"
                                onChange={handleatrComments}
                                value={comments === null ? "" : comments}
                                readOnly={crtAtrObj.strAtrStatus !== "Action Taken"}
                                style={{
                                  borderColor: reqCommentsBorderColor,
                                  backgroundColor: crtAtrObj.strAtrStatus === "Action Taken" ? "#FFFFFF" : "#EBEBE4"
                                }}
                              />
                              {reqCommentsError && (
                                <div className="inCorrectFileError">
                                  {reqCommentsError}
                                </div>
                              )}
                              <Hint direction={"end"}>
                                {comments.length} / {max}
                              </Hint>
                            </div>
                          </FieldWrapper>
                        </div>
                      )}

                      <div className="col-md-6">
                        <FieldWrapper>
                          <div className="k-form-field-wrap">
                            <label htmlFor="actionTaken">
                              Action Taken:
                              {isAssignee && isReviewer && (
                                <span className="required-asterisk">*</span>
                              )}
                            </label>
                            <TextArea
                              component={Input}
                              maxLength={maxActionTaken}
                              id="actionTaken"
                              name="ActionTaken"
                              value={actionTaken === null ? "" : actionTaken}
                              defaultValue={actionTaken}
                              style={{ borderColor: actionTakenBorderColor }}
                              onChange={handleOnChangeActionTaken}
                              readOnly={
                                !["Submitted", "Returned", "Pending ATR Review"].includes(
                                  crtAtrObj.strAtrStatus
                                )
                              }
                            />
                          </div>
                          <div className="_ATRFieldErrorMsg">
                            {actionTakenError}
                          </div>
                          <Hint direction={"end"}>
                            {actionTaken.length} / {maxActionTaken}
                          </Hint>
                        </FieldWrapper>
                      </div>

                      <div className="col-md-6">
                        <FieldWrapper>
                          <div className="k-form-field-wrap">
                            <label htmlFor="actionTakendate">
                              Action Taken Date:
                              {isAssignee && (
                                <span className="required-asterisk">*</span>
                              )}
                            </label>
                            {isAssignee ? (
                              <DatePicker
                                format={"dd-MM-yyyy"}
                                placeholder="Choose a date..."
                                onChange={handleOnChangeActionTakenDt}
                                value={actionTakenDt}
                                defaultValue={new Date(actionTakenDt)}
                                style={{ borderColor: actionTakenDateBorderColor, }}
                                readOnly={true}
                              />
                            ) : (
                              <Field
                                component={Input}
                                id="actionTkn"
                                name="actionTakenDate"
                                key={actionTakenDt}
                                defaultValue={new DateObject(
                                  new Date(actionTakenDt)
                                ).format("DD-MM-YYYY")}
                                readOnly={!isEditMode}
                              />
                            )}
                          </div>
                          <div className="_ATRFieldErrorMsg">
                            {actionTakenDateError}
                          </div>
                        </FieldWrapper>
                      </div>

                      {isRequester && (
                        <div className="col-md-6">
                          <FieldWrapper>
                            <div className="k-form-field-wrap">
                              <label htmlFor="comments">Comments:</label>
                              <TextArea
                                component={Input}
                                id="comments"
                                name="Comments"
                                value={comments}
                                defaultValue={comments}
                                onChange={handleatrComments}
                                readOnly={crtAtrObj.strAtrStatus !== "Action Taken"}
                                style={{
                                  borderColor: reqCommentsBorderColor,
                                  backgroundColor: crtAtrObj.strAtrStatus === "Action Taken" ? "#FFFFFF" : "#EBEBE4"
                                }}
                              />
                            </div>
                            <div className="_ATRFieldErrorMsg">
                              {reqCommentsError}
                            </div>
                          </FieldWrapper>
                        </div>
                      )}
                      {reviewerList.length > 0 && !externalType && (
                        <div className="col-md-6">
                          <FieldWrapper>
                            <div className="k-form-field-wrap">
                              {(crtAtrObj.strAtrStatus === "Returned" ||
                                (defaultAssignee && atrReviewersDTO.length === 0 &&
                                  crtAtrObj.defaultAssignee === accounts[0].username &&
                                  reviewerList.length > 0)) && (
                                  <div>
                                    <label htmlFor="actionTakendate">
                                      Reviewer:<span className="required-asterisk">*</span>
                                    </label>

                                    <DropDownList
                                      data={reviewerList.map((x) => x.approverEmailName)}
                                      onChange={handleReviewerChange}
                                      name="reviewer"
                                      style={{ borderColor: reviewerBorderColor }}
                                      value={
                                        crtAtrObj.strAtrStatus === "Returned" && selectedApproverName === ""
                                          ? atrReviewersDTO[0]?.atrReviewerEmailName
                                          : selectedApproverName
                                      }
                                    />
                                  </div>
                                )}
                            </div>
                          </FieldWrapper>
                        </div>
                      )}

                    </div>
                  </fieldset>

                  {isAssignee && (
                    <div>
                      <div className="SectionHeads eNoteAtrFormHdr row">
                        Attach Support Documents
                      </div>

                      {/* multidoc controller */}
                      <div className="SectionRow row">
                        <div className="Attachemntfileinfo-ind">
                          <input
                            type="file"
                            id="multiDoc"
                            multiple="multiple"
                            // onChange={(event) => multipleDocUpload(event, 'SupportingDocument')}
                            onChange={multipleDocUpload}
                            className="_multipledocUploadWidth"
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
                  )}

                  {!isAssignee && (
                    <div>
                      <div className="SectionHeads eNoteAtrFormHdr row">
                        Workflow Log
                      </div>
                      <div className="SectionRow row">
                        <table className="SectionRow cstATRFormTbl tableStyle">
                          <tr>
                            <th>Action</th>
                            <th>Action By</th>
                            <th>Action Date</th>
                          </tr>
                          {crtAtrObj.atrWorkflowsDTO?.map((obj, ind) => (
                            <tr key={ind}>
                              <td>{obj.action}</td>
                              {/* Bug fix - 293 - 27/03 */}
                              <td>{obj.actionByName}</td>
                              <td>
                                {new DateObject(new Date(obj.createdDate)).format(
                                  "DD-MMM-YYYY hh:mm A" // change -23/05 Seconds hand  was removed
                                )}
                              </td>
                            </tr>
                          ))}
                        </table>
                      </div>
                    </div>
                  )}

                  {!isAssignee && (
                    <div>
                      <div className="SectionHeads eNoteAtrFormHdr row">
                        File Attachments
                      </div>
                      <div className="SectionRow row   ">
                        <div className="table-responsive">
                          <table className="cstATRFormTbl tableStyle ">
                            <tr>
                              <th>Document link</th>
                              <th>Attached By</th>
                              <th>Attached Date</th>
                            </tr>
                            {crtAtrObj.atrSupportDocumentsDTO?.map((doc, ind) => (
                              <tr key={ind}>
                                <td>
                                  <span
                                    className="link-style"
                                    onClick={() =>
                                      getSupportDocHyperlink(
                                        doc.supportDocumentPath,
                                        doc.supportingDocumentFileName
                                      )
                                    }
                                  >
                                    {doc.supportingDocumentFileName}
                                  </span>
                                </td>
                                {/* Bug fix - 293 - 27/03 */}
                                <td>{doc.createdByName}</td>
                                <td>
                                  {new DateObject(
                                    new Date(doc.createdDate)
                                  ).format(
                                    "DD-MMM-YYYY hh:mm A" // change -23/05 Seconds hand  was removed
                                  )}
                                </td>
                              </tr>
                            ))}
                          </table>
                        </div>
                      </div>
                    </div>
                  )}

                  <div className="k-form-buttons _atrFormActionBtn">
                    {isAssignee && (
                      <Button
                        type={"submit"}
                        onClick={handleBtnActionComplete} disabled={actionInProgress}
                        className="k-button k-button-md k-rounded-md k-button-solid k-button-solid-base formBtnColor"
                      >
                        <span className="k-icon-sm k-icon k-font-icon k-i-track-changes-accept cursor atrFormBtnStyles"></span>
                        Action Completed
                      </Button>
                    )}

                    {isRequester && (
                      <Button
                        type={"submit"}
                        onClick={handleBtnAccept} disabled={actionInProgress}
                        className="k-button k-button-md k-rounded-md k-button-solid k-button-solid-base formBtnColor"
                      >
                        <span className="k-icon-xs k-icon k-font-icon k-i-checkmark cursor atrFormBtnStyles"></span>
                        {crtAtrObj.strAtrStatus === "Pending ATR Review" ? "Review" : "Accept"}
                      </Button>
                    )}
                    {isRequester && (
                      <Button
                        onClick={handleBtnReturn} disabled={actionInProgress}
                        type={"submit"}
                        className="k-button k-button-md k-rounded-md k-button-solid k-button-solid-base formBtnColor"
                      >
                        <span className="k-icon-xs k-icon k-font-icon  k-i-undo cursor atrFormBtnStyles"></span>
                        Return
                      </Button>
                    )}

                    <Button
                      type={"submit"}
                      onClick={handleExit}
                      className="k-button k-button-md k-rounded-md k-button-solid k-button-solid-base"
                    >
                      <span className="k-icon-sm k-icon k-font-icon  k-i-x-circle cursor atrFormBtnStyles"></span>
                      Exit
                    </Button>
                  </div>
                  {/* {isLoading && <PageLoader />} */}
                  {dialogOpen && (
                    <Dialog
                      title={<CustomConfirmDialogTitleBar />}
                      onClose={handleCloseDlg}
                    >
                      <p className="dialogcontent_">
                        {confirmDailogObj.Confirmtext}
                      </p>
                      <p className="dialogcontent_">
                        {confirmDailogObj.Description}
                      </p>
                      <DialogActionsBar>

                        {isAssignee && (
                          <Button
                            className="k-button k-button-md k-rounded-md k-button-solid k-button-solid-base formBtnColor"
                            onClick={handleDlgActionComplete}
                          >
                            <span className="k-icon k-font-icon  k-i-checkmark-circle cursor allIconsforPrimary-btn"></span>
                            Confirm
                          </Button>
                        )}

                        {isRequester && (
                          <Button
                            className="k-button k-button-md k-rounded-md k-button-solid k-button-solid-base formBtnColor"
                            onClick={handleDlgRequester}
                          >
                            <span className="k-icon k-font-icon  k-i-checkmark-circle cursor allIconsforPrimary-btn"></span>
                            Confirm
                          </Button>
                        )}
                        <Button
                          className="k-button k-button-md k-rounded-md k-button-solid k-button-solid-base"
                          onClick={handleCloseDlg}
                        >
                          <span className="k-icon k-font-icon  k-i-close-circle cursor allIconsforPrimary-btn"></span>
                          Cancel
                        </Button>
                      </DialogActionsBar>
                    </Dialog>
                  )}

                  {dialogSuccess && (
                    <Dialog
                      title={<CustomDialogTitleBar />}
                      onClose={handleClsDlgSuccess}
                    >
                      <p className="dialogcontent_">
                        ATR updation completed successfully!
                      </p>
                      <DialogActionsBar>
                        <Button
                          className="notifyDailogOkBtn k-button k-button-md k-rounded-md k-button-solid k-button-solid-base"
                          onClick={handleDlgBtnSuccess}
                        >
                          <span className="k-icon k-font-icon  k-i-redo cursor allIconsforPrimary-btn"></span>
                          Ok
                        </Button>
                      </DialogActionsBar>
                    </Dialog>
                  )}
                  {reviewerSuccessDialog && (
                    <Dialog
                      title={<CustomDialogTitleBar />}
                      onClose={() => setReviewerSuccessDialog(false)}
                    >
                      <p className="dialogcontent_">
                        {reviewerStatus}
                      </p>
                      <DialogActionsBar>
                        <Button
                          className="notifyDailogOkBtn k-button k-button-md k-rounded-md k-button-solid k-button-solid-base"
                          onClick={() => setReviewerSuccessDialog(false)}
                        >
                          <span className="k-icon k-font-icon  k-i-redo cursor allIconsforPrimary-btn"></span>
                          Ok
                        </Button>
                      </DialogActionsBar>
                    </Dialog>
                  )}
                  {dialogFailure && (
                    <Dialog
                      title={<CustomDialogTitleBar />}
                      onClose={handleClsDlgFailure}
                    >
                      <p className="dialogcontent_">{failureDailogTxt}</p>
                      <DialogActionsBar>
                        <Button
                          className="notifyDailogOkBtn k-button k-button-md k-rounded-md k-button-solid k-button-solid-base"
                          onClick={handleDlgBtnFailure}
                        >
                          <span className="k-icon k-font-icon  k-i-redo cursor allIconsforPrimary-btn"></span>
                          Ok
                        </Button>
                      </DialogActionsBar>
                    </Dialog>
                  )}
                  {/* passcode verification dialog*/}
                  {passcodeVerification && (
                    <Dialog
                      title="Passcode Verification"
                      className="mobile_passcode"
                      onClose={handlepasscodeClose}
                    >
                      <form
                        className="form-container dialogcontent_"
                      >
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
                            // Commented for passcode text field numeric input type --> Kavya(16-08)
                            // inputMode="numeric"
                            />
                            <span
                              className="eye-icon-view"

                              tabIndex="0" // Makes the span focusable for keyboard events
                              onMouseDown={() => setIsPasscodeVisible(true)}    // Show on mouse down
                              onMouseUp={() => setIsPasscodeVisible(false)}     // Hide on mouse up
                              onKeyDown={(e) => handleKeyDown(e)}                   // Show on key down (Enter or Space)
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
                            {error && <p className="error">{error}</p>}
                          </div>
                        </label>
                      </form>
                      <DialogActionsBar>
                        <Button
                          className="formBtnColor"
                          onClick={passcodeVerificationFunction}

                        >
                          <span className="k-icon k-font-icon  k-i-launch cursor allIconsforPrimary-btn"></span>
                          Verify
                        </Button>
                        <Button onClick={handlepasscodeClose}>
                          <span className="k-icon k-font-icon  k-i-close-circle cursor allIconsforPrimary-btn"></span>
                          Cancel
                        </Button>
                      </DialogActionsBar>
                    </Dialog>
                  )}
                  {/* Valid passcode dialog*/}
                  {validPasscode && (
                    <Dialog
                      title="Passcode Verification"
                      onClose={() => setValidPasscode(false)}
                    >
                      <div className="dialogcontent_">
                        <p>{validMsg}</p>
                      </div>
                      <DialogActionsBar>
                        <Button
                          className="formBtnColor"
                          onClick={handleRedirectToPasscode}
                        >
                          <span className="k-icon k-font-icon  k-i-launch cursor allIconsforPrimary-btn"></span>
                          Create Passcode
                        </Button>
                        <Button onClick={() => setValidPasscode(false)}>
                          <span className="k-icon k-font-icon  k-i-close-circle cursor allIconsforPrimary-btn"></span>
                          Cancel
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
    </>
  );
};