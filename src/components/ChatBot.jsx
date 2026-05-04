import React, { useState, useRef } from "react";
import "../styles/chatbot.css";
import { useMsal, useAccount } from "@azure/msal-react";
import { loginRequest } from "../config";
import { getAccessToken } from "../App";
import { Link } from "react-router-dom";

/* ================= ICONS ================= */
const SendIcon = () => (
  <svg width="20" height="20" viewBox="0 0 24 24" fill="white">
    <path d="M2 21L23 12L2 3V10L17 12L2 14V21Z" />
  </svg>
);

const StopIcon = () => (
  <svg width="20" height="20" viewBox="0 0 24 24" fill="white">
    <rect x="6" y="6" width="12" height="12" rx="2" />
  </svg>
);

// New Icons for Action Buttons
const SummaryIcon = () => (
  <svg width="18" height="18" viewBox="0 0 24 24" fill="currentColor" className="action-icon">
    <path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-5 14H7v-2h7v2zm3-4H7v-2h10v2zm0-4H7V7h10v2z"/>
  </svg>
);

const ApproveIcon = () => (
  <svg width="18" height="18" viewBox="0 0 24 24" fill="currentColor" className="action-icon">
    <path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/>
  </svg>
);

const RejectIcon = () => (
  <svg width="18" height="18" viewBox="0 0 24 24" fill="currentColor" className="action-icon">
    <path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/>
  </svg>
);

/* ================= PROMPT STATUS MAPPING ================= */
const promptStatusMap = [
  // NOTES
  { text: "notes pending for my attention", type: "NOTES", status: 3 },
  { text: "notes returned by me", type: "NOTES", status: 4 },
  { text: "notes approved by me", type: "NOTES", status: 5 },
  { text: "notes recommended by me", type: "NOTES", status: 9 },
  
  // ATRS
  { text: "atrs pending for my attention", type: "ATRS", atrStatus: 1 },
  { text: "atrs completed by me", type: "ATRS", atrStatus: 4 },
  { text: "atrs pending for my approval", type: "ATRS", atrStatus: 2 },
  { text: "atrs approved by me", type: "ATRS", atrStatus: 4 },
  { text: "atrs returned by me", type: "ATRS", atrStatus: 3 },
];

/* ================= COMPONENT ================= */
export const ChatBot = () => {
  const [messages, setMessages] = useState([
    {
      sender: "bot",
      text: "Hello! 👋 I can help you with your notes. How may I assist you today?",
    },
  ]);

  const [input, setInput] = useState("");
  const [isOpen, setIsOpen] = useState(false);
  const [isResponding, setIsResponding] = useState(false);

  const [isSummaryOpen, setIsSummaryOpen] = useState(false);
  const [summaryNoteNumber, setSummaryNoteNumber] = useState(null);
  const [summaryText, setSummaryText] = useState("");

  // New states for approve/reject functionality
  const [isApproveConfirmOpen, setIsApproveConfirmOpen] = useState(false);
  const [isRejectConfirmOpen, setIsRejectConfirmOpen] = useState(false);
  const [approveComment, setApproveComment] = useState(""); // New state for approve comment
  const [rejectComment, setRejectComment] = useState("");
  const [currentNote, setCurrentNote] = useState(null);
  const [isCommentInputOpen, setIsCommentInputOpen] = useState(false);
  const [commentModalType, setCommentModalType] = useState(""); // "approve" or "reject"
  const [isActionSuccessOpen, setIsActionSuccessOpen] = useState(false);
  const [actionMessage, setActionMessage] = useState("");
  const [isValidationOpen, setIsValidationOpen] = useState(false);
  const [validationMessage, setValidationMessage] = useState("");

  const [approverDetails, setApproverDetails] = useState(null);
  const [actionType, setActionType] = useState(""); // "approve" | "reject"

  const botTimeoutRef = useRef(null);

  const { accounts, instance } = useMsal();
  const account = useAccount(accounts[0] || {});
  const sessionId = accounts[0]?.homeAccountId;

  console.log("Session ID is: ", sessionId);
  
  /* ================= PREVIOUS BUTTON PROMPTS ================= */
  const promptOptions = [
    "My pending notes", 
    "My approved notes", 
    "My pending ATRs",
  ];

  /* ================= INTENT RESOLVER ================= */
  const resolveIntent = (question) => {
    if (typeof question !== "string") return { type: "UNKNOWN" };
    const q = question.toLowerCase();

    const matched = promptStatusMap.find((p) => q.includes(p.text.toLowerCase()));
    if (matched) return { type: "CUSTOM_PROMPT", data: matched };

    const normalized = q
      .replace(/pendng|pendin|oending/g, "pending")
      .replace(/aproved|aprovd/g, "approved")
      .replace(/atre|atrss/g, "atrs")
      .replace(/summery|summry|sumary/g, "summary")
      .replace(/not\s+id/g, "note id")
      .replace(/recommnded|recomended|reccommended/g, "recommended");

    const summaryMatch = normalized.match(/summary.*note\s*id\s*(\d+)/);

    if (summaryMatch) return { type: "SUMMARY", noteId: summaryMatch[1] };

    if (normalized.includes("pending") && normalized.includes("note")) return { type: "PENDING_NOTES" };
    if (normalized.includes("approved") && normalized.includes("note")) return { type: "APPROVED_NOTES" };
    if (normalized.includes("pending") && normalized.includes("atr")) return { type: "PENDING_ATRS" };
    
    // Handle "notes recommended by me" typed by user
    if (normalized.includes("recommended") && normalized.includes("note")) {
      return { 
        type: "CUSTOM_PROMPT", 
        data: { text: "notes recommended by me", type: "NOTES", status: 9 } 
      };
    }

    return { type: "UNKNOWN" };
  };

  /* ================= HELPER FUNCTIONS ================= */
  // Helper function to parse Python-style dictionary strings
 const parsePythonDictString = (data) => {
  if (typeof data !== 'string') return data;
  
  // Check if this is a summary string (starts with "Document Summary")
  if (data.trim().startsWith("Document Summary")) {
    return data; // Return the string as-is for summaries
  }
  
  try {
    // Handle None values
    data = data.replace(/None/g, 'null');
    
    // Handle single quotes
    data = data.replace(/'/g, '"');
    
    // Handle Python boolean values
    data = data.replace(/True/g, 'true').replace(/False/g, 'false');
    
    return JSON.parse(data);
  } catch (error) {
    console.error("Error parsing Python dict string:", error);
    // Return the original string if parsing fails (could be a summary)
    return data;
  }
};

  /* ================= API ================= */
  const fetchBotResponse = async (text) => {
    try {
      const accessToken = await getAccessToken({ ...loginRequest, account }, instance);

      const response = await fetch(
        "https://indianbank-webapp-aiteam-ceezd6dghdaheqe0.canadacentral-01.azurewebsites.net/process_query",
        {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            user_query: text,
            user_email: account?.username,
            bearer_token: accessToken,
            session_id: sessionId,
          }),
        }
      );

      const result = await response.json();
      
      // If data is a string, try to parse it
      if (typeof result.data === 'string') {
        try {
          // Parse the Python-style dictionary string
          result.data = parsePythonDictString(result.data);
        } catch (parseError) {
          console.error("Failed to parse data string:", parseError);
          // If parsing fails, set to empty object
          result.data = {};
        }
      }
      
      console.log("Parsed API Response:", result);
      return result;
    } catch (error) {
      console.error("Chatbot API error:", error);
      return {};
    }
  };

  /* ================= HELPERS ================= */
const filterNotesByStatus = (list, status) => {
  if (!Array.isArray(list)) {
    console.log("filterNotesByStatus: list is not an array:", list);
    return [];
  }

  // Convert status to array if single value is passed
  const statusList = Array.isArray(status) ? status : [status];

  return list.filter((item) => {
    if (!item) return false;

    const noteStatus = item.status ?? item.noteStatus ?? item.note_status;
    if (noteStatus === undefined || noteStatus === null) return false;

    return statusList.includes(Number(noteStatus));
  });
};


  const filterATRsByStatus = (list, atrStatus) => {
    if (!Array.isArray(list)) return [];
    return list.filter((item) => {
      if (!item) return false;
      
      const statusValue = item.atrStatus ?? item.status ?? item.atr_status;
      if (statusValue === undefined || statusValue === null) return false;
      
      return Number(statusValue) === Number(atrStatus);
    });
  };

  /* ================= HELPER FUNCTIONS ================= */
 // New helper to filter approver details for current user
const filterApproverForCurrentUser = (approverData, currentUserEmail) => {
  if (!Array.isArray(approverData)) {
    // Try to parse if it's a string
    try {
      approverData = parsePythonDictString(approverData);
      if (!Array.isArray(approverData)) return null;
    } catch (error) {
      console.error("Error parsing approver data:", error);
      return null;
    }
  }
  
  // Find approver with matching email
  const approver = approverData.find(approver => 
    approver.approverEmail?.toLowerCase() === currentUserEmail?.toLowerCase()
  );
  
  // Log approver status for debugging
  if (approver) {
    console.log("Approver Status:", approver.strApproverStatus);
    console.log("Approver Details:", approver);
  }
  
  return approver;
};

  // Handle Approve button click
// Handle Approve button click
const handleApproveClick = async (note) => {
  setCurrentNote(note);
  setActionType("approve");

  const accessToken = await getAccessToken({ ...loginRequest, account }, instance);
  const currentUserEmail = account?.username?.toLowerCase();

  const payload = {
    user_query: `approve note ${note.noteId}`,
    user_email: account?.username,
    bearer_token: accessToken,
    session_id: sessionId,
    approval_comments: "",
    rejection_comments: ""
  };

  const res = await fetch(
    "https://indianbank-webapp-aiteam-ceezd6dghdaheqe0.canadacentral-01.azurewebsites.net/process_query",
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    }
  );

  const result = await res.json();
  
  // Filter approver details for current user only
  const currentUserApprover = filterApproverForCurrentUser(result?.data, currentUserEmail);
  
  // if (!currentUserApprover) {
  //   setValidationMessage("You are not authorized to approve this note.");
  //   setIsValidationOpen(true);
  //   return;
  // }

  // Check if note is already approved by current user
  if (currentUserApprover.strApproverStatus?.toLowerCase() === "approved") {
    setValidationMessage(`Note ${note.noteNumber} is already approved.`);
    setIsValidationOpen(true);
    return;
  }

  // Check if note is already rejected by current user
  if (currentUserApprover.strApproverStatus?.toLowerCase() === "rejected") {
    setValidationMessage(`Note ${note.noteNumber} is already rejected. Cannot approve a rejected note.`);
    setIsValidationOpen(true);
    return;
  }

  setApproverDetails(currentUserApprover);
  
  // Open comment input modal for approve (comment is optional)
  setCommentModalType("approve");
  setIsCommentInputOpen(true);
};

  // Handle Reject button click
// Handle Reject button click
const handleRejectClick = async (note) => {
  setCurrentNote(note);
  setActionType("reject");

  const accessToken = await getAccessToken({ ...loginRequest, account }, instance);
  const currentUserEmail = account?.username?.toLowerCase();

  const payload = {
    user_query: `reject note ${note.noteId}`,
    user_email: account?.username,
    bearer_token: accessToken,
    session_id: sessionId,
    approval_comments: "",
    rejection_comments: ""
  };

  const res = await fetch(
    "https://indianbank-webapp-aiteam-ceezd6dghdaheqe0.canadacentral-01.azurewebsites.net/process_query",
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    }
  );

  const result = await res.json();
  
  // Filter approver details for current user only
  const currentUserApprover = filterApproverForCurrentUser(result?.data, currentUserEmail);
  
  // if (!currentUserApprover) {
  //   setValidationMessage("You are not authorized to reject this note.");
  //   setIsValidationOpen(true);
  //   return;
  // }

  // Check if note is already rejected by current user
  if (currentUserApprover.strApproverStatus?.toLowerCase() === "rejected") {
    setValidationMessage(`Note ${note.noteNumber} is already rejected.`);
    setIsValidationOpen(true);
    return;
  }

  // Check if note is already approved by current user
  if (currentUserApprover.strApproverStatus?.toLowerCase() === "approved") {
    setValidationMessage(`Note ${note.noteNumber} is already approved. Cannot reject an approved note.`);
    setIsValidationOpen(true);
    return;
  }

  setApproverDetails(currentUserApprover);
  
  // Open comment input modal for reject (comment is mandatory)
  setCommentModalType("reject");
  setIsCommentInputOpen(true);
};

  // Handle comment submission from modal
  const handleCommentSubmit = () => {
    if (commentModalType === "reject" && !rejectComment.trim()) {
      // For reject, comment is mandatory
      setValidationMessage("Please enter a comment before rejecting.");
      setIsValidationOpen(true);
      return;
    }
    
    // For approve, comment is optional (can be empty)
    if (commentModalType === "approve") {
      setIsCommentInputOpen(false);
      setIsApproveConfirmOpen(true);
    } else {
      setIsCommentInputOpen(false);
      setIsRejectConfirmOpen(true);
    }
  };

// Confirm Approve action
const confirmApprove = async () => {
  const loginUser = account?.username?.toLowerCase();
  
  // Double-check the approver email matches current user
  // if (loginUser !== approverDetails?.approverEmail?.toLowerCase()) {
  //   setIsApproveConfirmOpen(false);
  //   setValidationMessage("You are not authorized to approve this note.");
  //   setIsValidationOpen(true);
  //   return;
  // }

  // Additional check to prevent double approval
  if (approverDetails?.strApproverStatus?.toLowerCase() === "approved") {
    setIsApproveConfirmOpen(false);
    setValidationMessage(`Note ${currentNote.noteNumber} is already approved.`);
    setIsValidationOpen(true);
    return;
  }

  // ✅ ALLOW approve - only current user's entry will be processed
  const accessToken = await getAccessToken({ ...loginRequest, account }, instance);

  const payload = {
    user_query: "approve",
    user_email: account?.username,
    bearer_token: accessToken,
    session_id: sessionId,
    approver_details: approverDetails, // Only current user's approver details
    approval_comments: approveComment || "", // Use entered comment or default
    rejection_comments: ""
  };

  try {
    const response = await fetch(
      "https://indianbank-webapp-aiteam-ceezd6dghdaheqe0.canadacentral-01.azurewebsites.net/process_query",
      {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      }
    );
    
    const result = await response.json();
    console.log("Approval response:", result);

    setIsApproveConfirmOpen(false);
    setApproveComment(""); // Reset comment
    setActionMessage(`Note ${currentNote.noteNumber} approved successfully.`);
    setIsActionSuccessOpen(true);

    // Refresh the pending notes list
    // You can trigger a refetch here if needed
  } catch (error) {
    console.error("Approval error:", error);
    setValidationMessage("Failed to approve note. Please try again.");
    setIsValidationOpen(true);
  }
};
  // Confirm Reject action
  const confirmReject = async () => {
    const loginUser = account?.username?.toLowerCase();
    
    // Double-check the approver email matches current user
    // if (loginUser !== approverDetails?.approverEmail?.toLowerCase()) {
    //   setIsRejectConfirmOpen(false);
    //   setValidationMessage("You are not authorized to reject this note.");
    //   setIsValidationOpen(true);
    //   return;
    // }

    // ✅ ALLOW reject - only current user's entry will be processed
    const accessToken = await getAccessToken({ ...loginRequest, account }, instance);

    const payload = {
      user_query: "reject",
      user_email: account?.username,
      bearer_token: accessToken,
      session_id: sessionId,
      approver_details: approverDetails, // Only current user's approver details
      approval_comments: "",
      rejection_comments: rejectComment
    };

    try {
      const response = await fetch(
        "https://indianbank-webapp-aiteam-ceezd6dghdaheqe0.canadacentral-01.azurewebsites.net/process_query",
        {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(payload),
        }
      );
      
      const result = await response.json();
      console.log("Rejection response:", result);

      setIsRejectConfirmOpen(false);
      setRejectComment("");
      setActionMessage(`Note ${currentNote.noteNumber} rejected successfully.`);
      setIsActionSuccessOpen(true);

      // Refresh the pending notes list
      // You can trigger a refetch here if needed
    } catch (error) {
      console.error("Rejection error:", error);
      setValidationMessage("Failed to reject note. Please try again.");
      setIsValidationOpen(true);
    }
  };

  /* ================= SUMMARY FETCH ================= */
const fetchSummaryByNoteId = async (noteId) => {
  const question = `summarize note ${noteId}`;

  setMessages((prev) => [...prev, { sender: "user", text: question }]);
  setIsResponding(true);

  const apiResponse = await fetchBotResponse(question);
  
  // Fix: Extract summary from the correct location in API response
  let summary = "Summary not available";
  
  if (apiResponse?.data) {
    // If data is a string (like "Document Summary\n..."), use it directly
    if (typeof apiResponse.data === 'string') {
      summary = apiResponse.data;
    } 
    // If data is an object with an answer property
    else if (apiResponse.data.answer) {
      summary = apiResponse.data.answer;
    }
    // If data is the summary string directly
    else if (typeof apiResponse.data === 'object') {
      // Try to get any text content from the object
      summary = JSON.stringify(apiResponse.data);
    }
  }

  // Clean up the summary text - remove "Document Summary\n" prefix if present
  if (summary.startsWith("Document Summary\n")) {
    summary = summary.replace("Document Summary\n", "");
  }

  setSummaryNoteNumber(noteId);
  setSummaryText(summary);
  setIsSummaryOpen(true);

  setMessages((prev) => [
    ...prev,
    {
      sender: "bot",
      text: (
        <>
          <b>Summary for Note ID {noteId}</b>
          <p style={{ marginTop: "8px" }}>{summary}</p>
        </>
      ),
    },
  ]);

  setIsResponding(false);
};

  /* ================= TABLE RENDERER ================= */
  const renderTable = (title, list, showSummary = true, isPendingNotes = false) => {
    console.log("renderTable called with:", { title, list, showSummary, isPendingNotes });
    
    if (!Array.isArray(list) || list.length === 0) return <p>No records found.</p>;

    return (
      <>
        <b>{title} ({list.length})</b>
        <table className="chatbot-table">
          <thead>
            <tr>
              <th>SI.No</th>
              <th>Note Number</th>
              {isPendingNotes ? <th>Action</th> : (showSummary && <th>Summary</th>)}
            </tr>
          </thead>
          <tbody>
            {list.map((item, idx) => {
              const rowId = item.noteId || item.atrId;
              const displayNumber = item.noteNumber || "N/A";

              return (
                <tr key={rowId || idx}>
                  <td>{idx + 1}</td>
                  <td>
                    <Link
                      to={`/enoteviewform/${rowId}`}
                      target="_blank"
                      className="_note_number"
                      style={{ color: "red" }}
                    >
                      {displayNumber}
                    </Link>
                  </td>

                  {isPendingNotes ? (
                    <td>
                      <div className="action-icons-container">
                        <button
                          className="action-icon-btn summary-btn"
                          onClick={() => fetchSummaryByNoteId(rowId)}
                          disabled={isResponding}
                          title="View Summary"
                        >
                          <SummaryIcon />
                        </button>
                        <button
                          className="action-icon-btn approve-btn"
                          onClick={() => handleApproveClick(item)}
                          disabled={isResponding}
                          title="Approve Note"
                        >
                          <ApproveIcon />
                        </button>
                        <button
                          className="action-icon-btn reject-btn"
                          onClick={() => handleRejectClick(item)}
                          disabled={isResponding}
                          title="Reject Note"
                        >
                          <RejectIcon />
                        </button>
                      </div>
                    </td>
                  ) : (
                    showSummary && (
                      <td>
                        <button
                          className="summary-btn"
                          onClick={() => fetchSummaryByNoteId(rowId)}
                          disabled={isResponding}
                        >
                          View Summary
                        </button>
                      </td>
                    )
                  )}
                </tr>
              );
            })}
          </tbody>
        </table>
      </>
    );
  };

  /* ================= BOT RESPONSE ================= */
  const getBotResponse = (question, apiResponse) => {
    const data = apiResponse?.data || {};
    const intent = resolveIntent(question);

    console.log("API Response Data:", data);
    console.log("Intent:", intent);
    console.log("pendingNoteList:", data.pendingNoteList);

    // All notes/ATRs are in single lists
    const allNotes = data.pendingNoteList || [];
    const allATRs = data.pendingATRsList || [];

    console.log("All Notes Array:", allNotes);
    console.log("All Notes Length:", allNotes.length);

    // Handle typed prompts
    if (intent.type === "CUSTOM_PROMPT") {
      const { data: promptData } = intent;

      if (promptData.type === "NOTES") {
        const filteredNotes = filterNotesByStatus(allNotes, promptData.status);
        console.log(`Filtered Notes for status ${promptData.status}:`, filteredNotes);
        
        // Check if this is pending notes (status 3) to show action icons
        const isPending = promptData.status === 3;
        return renderTable(`Notes (${promptData.text})`, filteredNotes, !isPending, isPending);
      }

      if (promptData.type === "ATRS") {
        const filteredATRs = filterATRsByStatus(allATRs, promptData.atrStatus);
        return renderTable(`ATRs (${promptData.text})`, filteredATRs, false, false);
      }
    }

    // Handle previous buttons
    // if (intent.type === "PENDING_NOTES") {
    //   const filteredNotes = filterNotesByStatus(allNotes, 3);
    //   console.log("Pending Notes (status 3):", filteredNotes);
    //   return renderTable("Pending Notes", filteredNotes, false, true); // isPendingNotes = true
    // }

    if (intent.type === "PENDING_NOTES") {
  const filteredNotes = filterNotesByStatus(allNotes, [2, 3]);
  console.log("Pending Notes (status 2 or 3):", filteredNotes);
  return renderTable("Pending Notes", filteredNotes, false, true);
}

    
    if (intent.type === "APPROVED_NOTES") {
      const filteredNotes = filterNotesByStatus(allNotes, 5);
      console.log("Approved Notes (status 5):", filteredNotes);
      return renderTable("Approved Notes", filteredNotes, true, false); // showSummary = true
    }
    
    if (intent.type === "PENDING_ATRS") {
      const filteredATRs = filterATRsByStatus(allATRs, 5);
      return renderTable("Pending ATRs", filteredATRs, false, false);
    }

   if (intent.type === "SUMMARY") {
  // Fix: Extract summary from the correct location
  let summary = "Summary not available";
  
  if (typeof data === 'string') {
    summary = data;
  } 
  else if (data?.answer) {
    summary = data.answer;
  }
  else if (data) {
    // Try to stringify if it's an object
    summary = JSON.stringify(data);
  }

  // Clean up the summary text
  if (summary.startsWith("Document Summary\n")) {
    summary = summary.replace("Document Summary\n", "");
  }

  setSummaryNoteNumber(intent.noteId);
  setSummaryText(summary);
  setIsSummaryOpen(true);
  return (
    <>
      <b>Summary for Note ID {intent.noteId}</b>
      <p style={{ marginTop: "8px" }}>{summary}</p>
    </>
  );
}
    return "Sorry, I didn't understand. Try asking about notes or summaries.";
  };

  /* ================= SEND ================= */
  const handleSend = async (customText) => {
    if (customText && typeof customText !== "string") customText = undefined;
    if (!customText && !input.trim()) return;
    if (isResponding) return;

    const textToSend = customText ?? input;

    setMessages((prev) => [...prev, { sender: "user", text: textToSend }]);
    setInput("");
    setIsResponding(true);

    const apiResponse = await fetchBotResponse(textToSend);
    
    console.log("Full API Response after parsing:", apiResponse);
    console.log("Data field type:", typeof apiResponse.data);
    console.log("Data field value:", apiResponse.data);
    
    const botResponse = getBotResponse(textToSend, apiResponse);
    setMessages((prev) => [...prev, { sender: "bot", text: botResponse }]);
    
    setIsResponding(false);
  };

  const stopResponse = () => {
    clearTimeout(botTimeoutRef.current);
    setIsResponding(false);
  };

  const startNewChat = () => {
    setMessages([
      {
        sender: "bot",
        text: "Hello! 👋 I can help you with your notes. How may I assist you today?",
      },
    ]);
    setInput("");
    setIsResponding(false);
  };

  /* ================= UI ================= */
  return (
    <>
      {!isOpen && (
        <button className="chatbot-toggle" onClick={() => setIsOpen(true)}>💬</button>
      )}

      <div className={`chatbot-panel ${isOpen ? "open" : ""}`}>
        <div className="chatbot-header">
          <h6>eNote Assistant</h6>
          <button className="chatbot-close-btn" onClick={() => setIsOpen(false)}>×</button>
        </div>

        <div className="chatbot-messages">
          <button className="chatbot-new-chat-btn" onClick={startNewChat}>➕ New Chat</button>

          {messages.map((msg, idx) => (
            <div key={idx} className={`chatbot-message ${msg.sender}`}>{msg.text}</div>
          ))}

          {/* DISPLAY BUTTON PROMPTS */}
          <div className="chatbot-prompts">
            {promptOptions.map((option, i) => (
              <button
                key={i}
                className="chatbot-prompt-btn"
                onClick={() => handleSend(option)}
                disabled={isResponding}
              >
                {option}
              </button>
            ))}
          </div>
        </div>

        <div className="chatbot-input-row">
          <input
            value={input}
            onChange={(e) => setInput(e.target.value)}
            placeholder="Ask eNote Assistant..."
            onKeyDown={(e) => e.key === "Enter" && handleSend()}
            disabled={isResponding}
          />
          <button className="send-btn" onClick={() => (isResponding ? stopResponse() : handleSend())}>
            {isResponding ? <StopIcon /> : <SendIcon />}
          </button>
        </div>

        {/* Summary Modal */}
        {isSummaryOpen && (
          <div className="summary-modal-overlay">
            <div className="summary-modal">
              <div className="summary-modal-header">
                <span>Note Summary</span>
                <button className="summary-close-btn" onClick={() => setIsSummaryOpen(false)}>×</button>
              </div>
              <div className="summary-modal-body">
                <p><b>Note ID:</b> {summaryNoteNumber}</p>
                <p style={{ marginTop: "10px" }}>{summaryText}</p>
              </div>
            </div>
          </div>
        )}

        {/* Approve Confirmation Modal */}
        {isApproveConfirmOpen && (
          <div className="summary-modal-overlay">
            <div className="summary-modal">
              <div className="summary-modal-header">
                <span>Confirm Approval</span>
                <button className="summary-close-btn" onClick={() => setIsApproveConfirmOpen(false)}>×</button>
              </div>
              <div className="summary-modal-body">
                <p>Are you sure you want to approve <b>{currentNote?.noteNumber}</b>?</p>
                {approveComment && (
                  <div className="comment-preview">
                    <p><b>Comment:</b> {approveComment}</p>
                  </div>
                )}
                <div className="modal-buttons">
                  <button className="cancel-btn" onClick={() => setIsApproveConfirmOpen(false)}>
                    Cancel
                  </button>
                  <button className="confirm-btn" onClick={confirmApprove}>
                    OK
                  </button>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* Reject Confirmation Modal */}
        {isRejectConfirmOpen && (
          <div className="summary-modal-overlay">
            <div className="summary-modal">
              <div className="summary-modal-header">
                <span>Confirm Rejection</span>
                <button className="summary-close-btn" onClick={() => setIsRejectConfirmOpen(false)}>×</button>
              </div>
              <div className="summary-modal-body">
                <p>Are you sure you want to reject <b>{currentNote?.noteNumber}</b> with this comment?</p>
                <div className="comment-preview">
                  <p><b>Comment:</b> {rejectComment}</p>
                </div>
                <div className="modal-buttons">
                  <button className="cancel-btn" onClick={() => setIsRejectConfirmOpen(false)}>
                    Cancel
                  </button>
                  <button className="confirm-btn" onClick={confirmReject}>
                    Confirm
                  </button>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* Comment Input Modal (used for both approve and reject) */}
        {isCommentInputOpen && (
          <div className="summary-modal-overlay">
            <div className="summary-modal">
              <div className="summary-modal-header">
                <span>{commentModalType === "approve" ? "Approve Note" : "Reject Note"}</span>
                <button className="summary-close-btn" onClick={() => {
                  setIsCommentInputOpen(false);
                  setApproveComment("");
                  setRejectComment("");
                }}>×</button>
              </div>
              <div className="summary-modal-body">
                <p>
                  {commentModalType === "approve" 
                    ? `Please provide a comment for approving ${currentNote?.noteNumber}:`
                    : `Please provide a comment for rejecting ${currentNote?.noteNumber}:`}
                </p>
                <textarea
                  className="comment-textarea"
                  value={commentModalType === "approve" ? approveComment : rejectComment}
                  onChange={(e) => {
                    if (commentModalType === "approve") {
                      setApproveComment(e.target.value);
                    } else {
                      setRejectComment(e.target.value);
                    }
                  }}
                  placeholder={commentModalType === "approve" 
                    ? "Enter your comment here..." 
                    : "Enter your comment here..."
                  }
                  rows="4"
                />
                <div className="modal-buttons">
                  <button className="cancel-btn" onClick={() => {
                    setIsCommentInputOpen(false);
                    setApproveComment("");
                    setRejectComment("");
                  }}>
                    Cancel
                  </button>
                  <button className="confirm-btn" onClick={handleCommentSubmit}>
                    {commentModalType === "approve" ? "Next" : "Submit"}
                  </button>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* Action Success Modal */}
        {isActionSuccessOpen && (
          <div className="summary-modal-overlay">
            <div className="summary-modal">
              <div className="summary-modal-header">
                <span>Success</span>
                <button className="summary-close-btn" onClick={() => setIsActionSuccessOpen(false)}>×</button>
              </div>
              <div className="summary-modal-body">
                <p>{actionMessage}</p>
                <div className="modal-buttons" style={{ justifyContent: "center" }}>
                  <button className="confirm-btn" onClick={() => setIsActionSuccessOpen(false)}>
                    OK
                  </button>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* Validation Error Modal */}
        {isValidationOpen && (
          <div className="summary-modal-overlay">
            <div className="summary-modal">
              <div className="summary-modal-header">
                <span>Validation Required</span>
                <button className="summary-close-btn" onClick={() => setIsValidationOpen(false)}>×</button>
              </div>
              <div className="summary-modal-body">
                <p>{validationMessage}</p>
                <div className="modal-buttons" style={{ justifyContent: "center" }}>
                  <button className="confirm-btn" onClick={() => setIsValidationOpen(false)}>
                    OK
                  </button>
                </div>
              </div>
            </div>
          </div>
        )}
      </div>
    </>
  );
};