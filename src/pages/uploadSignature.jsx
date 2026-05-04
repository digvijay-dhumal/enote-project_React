import React, { useState, useEffect } from "react";
import { Button } from "@progress/kendo-react-buttons";
import { Dialog, DialogActionsBar } from "@progress/kendo-react-dialogs";
import { MultiColumnComboBox } from "@progress/kendo-react-dropdowns";

import Navbar from "../components/navbar";
import Footer from "../components/footer";
import { Sidebar } from "../components/sidebar";
import { MenuNavContainer } from "../components/menu";

import { getAccessToken } from "../App";
import { useMsal, useAccount } from "@azure/msal-react";
import {
  API_COMMON_HEADERS,
  API_BASE_URL,
  API_ENDPOINTS,
  loginRequest,
} from "../config";

import "../styles/passcode.css";

export const UploadSignature = () => {
  const [file, setFile] = useState(null);
  const [preview, setPreview] = useState(null);
  const [findAdminOrDepartmentAdmin, setFindAdminOrDepartmentAdmin] =
    useState(null);

  const [orgEmployees, setOrgEmployees] = useState([]);
  const [selectedApprover, setSelectedApprover] = useState(null);
  const [isMobile] = useState(window.innerWidth < 768);

  const [dialogOpen, setDialogOpen] = useState(false);
  const [dialogMessage, setDialogMessage] = useState("");

  const { accounts, instance } = useMsal();
  const account = useAccount(accounts[0] || {});

  useEffect(() => {
    loadUserAccessDetails();
  }, []);

  const orgUsersPplPickerColumnMapping = [
    { field: "displayName", header: "Person Name", width: "200px" },
    { field: "department", header: "Department", width: "180px" },
    { field: "jobTitle", header: "Designation", width: "180px" },
  ];

  const mobileColumns = [
    { field: "displayName", header: "Person Name", width: "120px" },
    { field: "department", header: "Department", width: "110px" },
    { field: "jobTitle", header: "Designation", width: "110px" },
  ];

  const loadUserAccessDetails = async () => {
    try {
      const accessToken = await getAccessToken(
        { ...loginRequest, account },
        instance
      );

      const response = await fetch(
        `${API_BASE_URL}${API_ENDPOINTS.Find_Super_Dept_Admin(
          accounts[0]?.username
        )}`,
        {
          method: "GET",
          headers: { ...API_COMMON_HEADERS, Authorization: `Bearer ${accessToken}` },
        }
      );

      const deptData = await response.json();
      console.log(deptData);
      setFindAdminOrDepartmentAdmin(deptData);
    } catch (error) {
      console.error("Error loading dept admin:", error);
    }
  };


const handleFileChange = (event) => {
  const selectedFile = event.target.files[0];
  if (!selectedFile) return;

  const fileName = selectedFile.name.toLowerCase();
  const allowedExtensions = [".jpg", ".jpeg"];

  const isValidExtension = allowedExtensions.some(ext =>
    fileName.endsWith(ext)
  );


  if (!isValidExtension) {
    setDialogMessage("Only JPG or JPEG files are allowed.");
    setDialogOpen(true);
    event.target.value = "";
    return;
  }

  // Size validation (50 kB)
const maxSize = 50 * 1024; // 50 KB

if (selectedFile.size > maxSize) {
  setDialogMessage("File size should not exceed 50 KB.");
  setDialogOpen(true);
  event.target.value = "";
  return;
}


  
  setFile(selectedFile);
  setPreview(URL.createObjectURL(selectedFile));
};



  // const handleFileChange = (event) => {
  //   const selectedFile = event.target.files[0];
  //   if (!selectedFile) return;

  //   const validTypes = ["image/jpeg", "image/jpg", "image/png"];
  //   if (!validTypes.includes(selectedFile.type)) {
  //     alert("Only JPG, JPEG, or PNG files are allowed.");
  //     return;
  //   }

  //   const maxSize = 1 * 1024 * 1024;
  //   if (selectedFile.size > maxSize) {
  //     alert("File size should not exceed 1 MB.");
  //     return;
  //   }

  //   setFile(selectedFile);
  //   setPreview(URL.createObjectURL(selectedFile));
  // };

  const handleOpenReviewer = () => {
    setOrgEmployees([]);
  };

  const handleApproverChange = (event) => {
    setSelectedApprover(event.value || null);
  };

  const pplFilterMultiColumn = async (event) => {
    const searchText =
      (event.filter && event.filter.value) ||
      (event.target && event.target.value) ||
      "";

    if (searchText.length < 4) {
      setOrgEmployees([]);
      return;
    }

    try {
      const accessToken = await getAccessToken({ ...loginRequest, account }, instance);

      const res = await fetch(
        `${API_BASE_URL}${API_ENDPOINTS.Search_UserDetails(encodeURIComponent(searchText))}`,
        {
          method: "GET",
          headers: { ...API_COMMON_HEADERS, Authorization: `Bearer ${accessToken}` },
        }
      );

      if (!res.ok) throw new Error(`HTTP error! Status: ${res.status}`);
      const data = await res.json();
      const usersArr = Array.isArray(data) ? data : [];

      const orgUsers = usersArr.map((x, idx) => ({
        displayName: x.displayName || "NA",
        department: x.department || "NA",
        jobTitle: x.jobTitle || "NA",
        mail: x.userPrincipalName || x.mail || "",
        srNo: x.srNo || idx + 1,
      }));

      setOrgEmployees(orgUsers);
    } catch (err) {
      console.error("User search error:", err);
      setOrgEmployees([]);
    }
  };

  const handleSubmit = async () => {
    if (!file) {
      setDialogMessage("Please select a signature image before submitting.");
      setDialogOpen(true);
      return;
    }

    // if (!selectedApprover) {
    //   setDialogMessage("Please select an approver.");
    //   setDialogOpen(true);
    //   return;
    // }

    if (!selectedApprover) {
  setDialogMessage("Please select an approver.");
  setDialogOpen(true);
  return;
}

// Same user validation
if (
  selectedApprover.mail?.toLowerCase() ===
  accounts[0]?.username?.toLowerCase()
) {
  setDialogMessage("You cannot select yourself as approver.");
  setDialogOpen(true);
  return;
}


    const base64 = await new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result.split(",")[1]);
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });

    const partSize = Math.ceil(base64.length / 10);
    const parts = Array.from({ length: 10 }, (_, i) =>
      base64.substring(i * partSize, (i + 1) * partSize)
    );

    const srNo = findAdminOrDepartmentAdmin?.srNo || "UnknownUser";
    const extension = file.name.split(".").pop();
    const fileName = `${srNo}.${extension}`;

    const payload = {
      srNo,
      userName: accounts[0]?.username,
      approverMail: selectedApprover.mail,
      signBlobPath1: parts[0],
      signBlobPath2: parts[1],
      signBlobPath3: parts[2],
      signBlobPath4: parts[3],
      signBlobPath5: parts[4],
      signBlobPath6: parts[5],
      signBlobPath7: parts[6],
      signBlobPath8: parts[7],
      signBlobPath9: parts[8],
      signBlobPath10: parts[9],
      fileName,
    };

    try {
      const accessToken = await getAccessToken({ ...loginRequest, account }, instance);

      const response = await fetch(`${API_BASE_URL}${API_ENDPOINTS.SignUpload}`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${accessToken}`,
        },
        body: JSON.stringify(payload),
      });

      if (response.ok) {
        setDialogMessage("Signature uploaded successfully!");
        setFile(null);
        setPreview(null);
        setSelectedApprover(null);
      } else {
        const err = await response.text();
        setDialogMessage("Your signature is already under approval. Please wait for the process to complete.");

      }
    } catch (error) {
      console.error(error);
      setDialogMessage("An error occurred while uploading signature.");
    }

    setDialogOpen(true);
  };

  return (
    <div>
      <Navbar header="IB Smart Office - eNote" />
      <Sidebar />
      <MenuNavContainer isMenuPage />

      <div className="container largeScreen text-center">
        <div className="SectionHeads row mobileSectionHeads">Upload Signature</div>

        <div className="card1 addPasscode1 p-4">
          <div className="form-container1">
            <div className="row g-4">
              {/* LEFT SIDE - FILE UPLOAD */}
              <div className="col-md-6" style={{margin: "0"}}>
                <div className="file-section p-3 d-flex flex-column h-100">
                  <div className="mb-3">
                    <label className="form-label fw-bold d-block mb-2">Choose File</label>
                    <div className="file-input-wrapper position-relative">
                      <input
                        type="file"
                        id="fileInput"
                        accept="image/jpeg, image/jpg"
                        onChange={handleFileChange}
                        className="form-control d-none"
                      />
                      <label
                        htmlFor="fileInput"
                        className="btn btn-outline-primary w-100 text-start"
                        style={{
                          padding: "8px 12px",
                          backgroundColor: "#f8f9fa",
                          cursor: "pointer",
                          height: "32px",
                          borderRadius: "5px",
                        }}
                      >
                        <span className="text-muted">
                          {file ? file.name : "No file chosen"}
                        </span>
                      </label>
                    </div>
                  </div>
                  <p 
                    className="mb-3" 
                    style={{ 
                      fontSize: "12px", 
                      lineHeight: "1.4",
                      textAlign: "left",
                      color: "white"
                    }}
                  >
                    Your signature image must be JPG or JPEG (Max 50KB)
                  </p>
                  
                  <div >
                    {preview && (
                      <div className="preview-container">
                        <img 
                          src={preview} 
                          alt="Signature Preview" 
                          style={{ 
                            maxWidth: "100px", 
                            maxHeight: "100px"
                          }} 
                        />
                      </div>
                    )}
                  </div>
                </div>
              </div>

              {/* RIGHT SIDE - APPROVER SELECTION */}
              <div className="col-md-6" style={{margin: "0"}}>
                <div className="approver-section p-3 d-flex flex-column h-100">
                  <div className="mb-3">
                    <label className="form-label fw-bold d-block mb-2">Select Approver</label>
                    <MultiColumnComboBox
                      data={orgEmployees}
                      dataItemKey="mail"
                      textField="displayName"
                      filterable
                      columns={isMobile ? mobileColumns : orgUsersPplPickerColumnMapping}
                      value={selectedApprover}
                      onFilterChange={pplFilterMultiColumn}
                      onChange={handleApproverChange}
                      onOpen={handleOpenReviewer}
                      placeholder="Enter Approver"
                      style={{ width: "100%" }}
                    />
                  </div>
                  <div>
                    <div className="text-dark small" style={{ fontSize: "12px", lineHeight: "1.2", textAlign: "left" }}>
                      {/* <p className="mb-1" style={{color: "white"}}>Select an approver from the list</p> */}
                      <p className="mb-0" style={{color: "white"}}>Type at least 4 characters to search</p>
                    </div>
                  </div>
                </div>
              </div>
            </div>

            {/* SUBMIT BUTTON - CENTERED BELOW */}
            <div className="row">
              <div className="col-12">
                <div className="text-center">
                  <Button
                    type="submit"
                    className="FormButtons k-button k-button-md k-rounded-md k-button-solid k-button-solid-base"
                    onClick={handleSubmit}
                    style={{
                      padding: "8px 30px",
                      fontSize: "16px"
                    }}
                  >
                    <span className="k-icon k-font-icon k-i-launch cursor allIconsforPrimary-btn me-2"></span>
                    <span>Submit</span>
                  </Button>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>

      <div className="pgFooterContainer">
        <Footer />
      </div>

      {dialogOpen && (
        <Dialog title="Alert!" onClose={() => setDialogOpen(false)} className="dialogcontent_Refer">
          <p className="dialogcontent_">{dialogMessage}</p>
          <DialogActionsBar>
            <div className="w-100 text-center">
              <Button onClick={() => setDialogOpen(false)} className="notifyDailogOkBtn k-button k-button-md k-rounded-md k-button-solid k-button-solid-base">
                <span className="k-icon k-font-icon k-i-redo cursor allIconsforPrimary-btn me-2"></span>
                <span>Ok</span>
              </Button>
            </div>
          </DialogActionsBar>
        </Dialog>
      )}
    </div>
  );
};