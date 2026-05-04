/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

import { LogLevel } from "@azure/msal-browser";

/**
 * Configuration object to be passed to MSAL instance on creation.
 * For a full list of MSAL.js configuration parameters, visit:
 * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/configuration.md
 */
// "Instance": "https://login.microsoftonline.com/",
// "Domain": "xencia.com",
// "TenantId": "75f2a99b-01fd-48f2-ac60-d4a7a44fd0cc",
// "ClientId": "e77908ab-8705-4f9a-b103-515ef6f69535",
// "CallbackPath": "/signin-oidc",
// "Scopes": "access_as_user"\

export const msalConfig = {
  auth: {
    // clientId: "b3fc558c-f399-4d3b-aa6c-ac70624e07a8",
    // authority: "https://login.microsoftonline.com/75f2a99b-01fd-48f2-ac60-d4a7a44fd0cc",
    // redirectUri: "https://inb-enote.xencia.com",
    clientId: process.env.REACT_APP_ClientId,
    authority: process.env.REACT_APP_Authority,
    redirectUri: process.env.REACT_APP_eNote_URL,
    // postLogoutRedirectUri: process.env.REACT_APP_Home_URL

    /*clientId: "ab1333a0-e6bc-4ba8-b67b-4f4378b37e7e",
    authority: "https://login.microsoftonline.com/19e8abd9-a4ec-4a12-a8e7-ad170593e0b2",
    // redirectUri: "https://smartofficeDev-enote.indianbank.in",
    redirectUri: "https://smartofficeuat-enote.indianbank.in",*/

    // clientId: "e77908ab-8705-4f9a-b103-515ef6f69535",
    // clientId: "73c2abaa-f13a-415c-aa1a-7cce420f85b0",
    // redirectUri: "https://inb.xencia.com",
    // redirectUri: "http://localhost:3000",
    // redirectUri: "https://20.55.197.131"
  },
  cache: {
    cacheLocation: "sessionStorage", // This configures where your cache will be stored
    storeAuthStateInCookie: true, // Set this to "true" if you are having issues on IE11 or Edge
  },
  system: {
    loggerOptions: {
      loggerCallback: (level, message, containsPii) => {
        if (containsPii) {
          return;
        }
        switch (level) {
          case LogLevel.Error:
            return;
          case LogLevel.Info:
            return;
          case LogLevel.Verbose:
            return;
          case LogLevel.Warning:
            return;
          default:
            return;
        }
      },
    },
  },
};

/**
 * Scopes you add here will be prompted for user consent during sign-in.
 * By default, MSAL.js will add OIDC scopes (openid, profile, email) to any login request.
 * For more information about OIDC scopes, visit:
 * https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-permissions-and-consent#openid-connect-scopes
 */
export const loginRequest = {
  // scopes: ['https://graph.microsoft.com/User.Read'],
  // AUD: "api://b3fc558c-f399-4d3b-aa6c-ac70624e07a8",
  scopes: [process.env.REACT_APP_Scopes],
  // scopes: ['User.Read', 'User.ReadBasic.All', 'email'],
  // scopes: ["api://b3fc558c-f399-4d3b-aa6c-ac70624e07a8/User.Read"],
  // scopes: ["api://ab1333a0-e6bc-4ba8-b67b-4f4378b37e7e/IBSmartOfficeScope"],
  // scopes: ["https://graph.microsoft.com/User.Read"],
  // loginHint: "" 
  // loginHint: "admin@M365x70282966.onmicrosoft.com" // or your login hint if needed
};

export const API_COMMON_HEADERS = {
  "Content-type": "application/json; charset=UTF-8",
  "Access-Control-Allow-Origin": process.env.REACT_APP_eNote_URL
}

/**
 * Add here the scopes to request when obtaining an access token for MS Graph API. For more information, see:
 * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/resources-and-scopes.md
 */
export const graphConfig = {
  graphMeEndpoint: "https://graph.microsoft.com/v1.0/me",
};

export const API_BASE_URL = process.env.REACT_APP_APIBase_URL;
export const IB_Domain_URL = process.env.REACT_APP_Home_URL;
export const IB_eNoteDomain_URL = process.env.REACT_APP_eNote_URL;
export const IB_eCommitteeDomain_URL = process.env.REACT_APP_eCommittee_URL;
export const IB_eDakDomain_URL = process.env.REACT_APP_eDak_URL;
export const IB_eReceptionDomain_URL = process.env.REACT_APP_eReception_URL;
export const IB_eArchivalDomain_URL = process.env.REACT_APP_eArchival_URL

export const chunkSize = 1024 * 1024 * 2; // 5MB
export const charValidation = /^[a-zA-Z0-9.,:;?!_@#₹ \t\n&"'\/\-\$%\(\)`]*$/;

export const sanitizeInput = (input) => {
  // Includes: LRM, RLM, PDF, LRE, RLE, LRO, RLO, FSI, PDI, ZWSP, NBSP
  return input.replace(/[\u200E\u200F\u202A-\u202E\u2066-\u2069\u200B\u00A0]/g, '').trim();
};

export const API_ENDPOINTS = {
  GET_DROPDOWNDATA: "/api/ENote/getDropdownData",
  GET_UserDetails: "/api/Azure/GetUserDetailsGraphALL",
  GET_UserDetails_Search: "/api/Azure/GetUserDetailsGraphALL",
  ATR_GetRequest: "/api/ATRStatus/Atrs",
  ATR_GetAllRequests: "/api/ATRStatus/GetAtrsStatusBaseAll",
  ATR_GetRequestsByStatus: "/api/ATRStatus/GetAtrsStatusBase",
  ATR_AssigneeAction: "/api/ATRStatus/AtrStatusUpdate",
  // ATR_AssigneeAction: "/api/ATRStatus/AtrStatusUpdateFileUpload",
  // ATR_AssigneeAction: "/api/ATRStatus/AtrStatusUpdateFileUpload",
  ATR_RequesterAction: "/api/ATRStatus/AtrStatusUpdateFromRequester",
  eNote_GetRequests: "/api/ENote/getNoteList",
  eNote_GetRequestsRoleBased: "/api/NavigationMenu/getNoteList",
  eNote_GetNotedNoteRequests: "/api/ENote/getNotedNoteList",
  eNote_GetEDMDNoteList: "/api/ENote/getEDMDNoteList",
  eNote_GetNoteSecretaryofApprover: "/api/NoteApproversMaster/GetNoteSecretaryfromApprover",
  // eNote_AddNote: "/api/ENote/AddNoteFileUpload",
  // eNote_EditNote: "/api/ENote/EditNoteFileUpload",
  eNote_AddNote: "/api/ENote/AddNote",
  eNote_EditNote: "/api/ENote/EditNote",
  eNote_GetGeneralDetails: "/api/ENote/getENoteGeneralDetails",
  eNote_ChangeApprover: "/api/ENote/ChangeApprover",
  // eNote_ChangeApprover: "/api/ENote/ChangeApproverFileUpload",
  eNote_GetNoteSecretary: "/api/NoteSecretary/GetNoteSecretary",
  eNote_CallBackNote: "/api/ENote/CallBackNote",
  eNote_CancelNote: "/api/ENote/CancelNote",
  eNote_NoteApproverStatusChange: "/api/ENote/NoteApproverStausChange",
  // eNote_NoteApproverStatusChange: "/api/ENote/NoteApproverStausChangeTesting",

  eNote_UpdateNoteReferresStatus: "/api/NoteReferre/UpdateNoteReferrerStatus",
  // eNote_UpdateNoteReferresStatus: "/api/NoteReferre/UpdateNoteReferrerStatusFileUpload",

  eNote_InserNoteSecretary: "/api/NoteSecretary/InsertNoteSecretary",
  // eNote_InserNoteSecretary: "/api/NoteSecretary/InsertNoteSecretaryFileUpload",

  eNote_AddNoteMarkedInformation: "/api/NoteMarkedInformation/AddNoteMarkedInformation",
  eNote_GetATRCreators: "/api/ATRCreators/getATRCreators",
  eNote_GetATRAssignees: "/api/ATRAssignees/getATRAssignees",
  eNote_SearchNotes: "/api/Search/SearchNotes",
  eNote_GetRefferedNoteList: "/api/ENote/getRefferedNoteList",
  eNote_findIsloginUserSecretary: "/api/NoteSecretary/IdentifySecretary",
  eNote_GetNoteChartData: "/api/ChartInfo/GetNoteChartInfo",
  eNote_GetNoteApprovers: (department) => `/api/NoteApproversMaster/getNoteApprovers/?department=${encodeURIComponent(department)}`,
  eNote_GetNoteReviewer: (department) => `/api/NoteApproversMaster/getNoteReviewer/?department=${encodeURIComponent(department)}`,
  GET_UserDetailsByPrincipalName: (username) => `/api/Azure/GetUserDetailsByPrincipalName?userPrincipalName=${username}`,
  Search_UserDetails: (username) => `/api/Azure/GetUserDetailsByStartChar?userPrincipalName=${username}`,
  Find_Super_Dept_Admin: (username) => `/api/Department/FindAdminorDepartmentAdmin?userPrincipalName=${username}`,
  GET_ReportChartsForAdmin: "/api/ChartInfo/GetReportPowerBi",
  GET_DeatilsForPending: "/api/ChartInfo/GetReportPowerBiPending",
  GET_AdminType: "/api/Admin/GetAdminType",
  GET_Departments: "/api/Department/getDepartments",
  Get_ATRReport: "/api/ChartInfo/GetATRReport",
  GET_Departments_List: "/api/ENote/getDepartments",
  GET_SupportingDoc: "/api/ENote/getEnoteGeneralDetailsSupportDoc",
  GET_ATRSupportingDoc: "/api/ATRStatus/getSupportFilesAtrs",
  GET_NotePdf: "/api/ENote/getEnoteGeneralDetailsNotePdf",
  GET_WordDoc: "/api/ENote/getEnoteGeneralDetailsWordPdf",
  GET_Base64: `/api/ENote/GetBase64`,
  // GET_Base64:`/api/ENote/GetBase64Chunks`,
  eNote_AddPasscode: "/api/NotePasscode/AddPasscode",
  eNote_SendOTP: "/api/NotePasscode/SendOTP",
  eNote_VerifyPasscode: "/api/NotePasscode/VerifyPasscode",
  eNote_VerifyUserPasscode: (username) => `/api/NotePasscode/PasscodeAvailability?userPrincipalName=${username}`,
  // GET_Base64PDF: `/api/ENote/getENotePDFBase64`,
  GET_Base64PDF:`/api/ENote/GetENotePDFRangeProcess`,
    // GET_Base64PDF:`/api/ENote/getENotePDFBase64`,
    // GET_Base64PDF:`/api/ENote/GetENotePDFBase64Chunk`,

  GET_NoteViewEligibility: "/api/ENote/getNoteViewEligibility",
  UploadChunks: "/api/FileUpload/UploadChunks",
  UploadComplete: "/api/FileUpload/UploadComplete",
  atrReviewerDropdown: "/api/ENote/getNoteReviewer",
  atrTypeDropdown: "/API/ENote/getdropdown",
  atrInternalDropdeown: "/api/ATRStatus/GetApproverReviewerList",
  eNote_findATRSecretaries: "/api/NoteSecretary/IdentifyATRSecretary",
  eNote_ATRBulkUpdate: "/api/ATRStatus/AtrBulkUpdate",
  eNote_uploadChunk: "/api/FileUpload/UploadChunks",
  eNote_uploadComplete: "/api/FileUpload/UploadComplete",
  eNote_GistUploadComplete: "/api/FileUpload/GistUploadComplete",
  eNote_ATRUploadComplete: "/api/FileUpload/AtrUploadComplete",
  eNote_ATRViewEligibility: "/api/AtrStatus/getAtrViewEligibility",
  UploadBase64Stream: "/api/FileUpload/UploadBase64Stream",
  UploadBase64ATR: "/api/FileUpload/UploadAtrAsBase64Stream",
  UploadBase64Gist: "/api/FileUpload/UploadGistAsBase64Stream",

  SaveApproverComment:"/api/ENote/SaveApproverComment",
  deleteComment:"/api/ENote/DeleteApproverComment",
  GET_DROPDOWNDATA_ereception: "/api/EReception/getEReceptionDropdownData",



  // DeleteDraft API
  eNote_DeleteDraft: "/api/ENote/DeleteDraft",

    //signature
  GET_APPROVER_PENDING_LIST:"/api/Admin/GetApproverPendingList",
  SignUpload: "/api/Admin/SignUpload",


   AddAnnotationPDF: "/api/AnnotationDetails/addAnnotation",
   updateAnnotationPDF: "/api/AnnotationDetails/updateAnnotation",
};