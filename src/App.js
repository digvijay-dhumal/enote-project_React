import { useEffect,createContext, useContext, useState } from "react";
import { Landingpage } from "./pages/landingpage";
import { Dialog, DialogActionsBar } from "@progress/kendo-react-dialogs";
import { Button } from "@progress/kendo-react-buttons";
import { BrowserRouter as Router, Routes, Route } from "react-router-dom";
import { ENoteHome } from "./pages/eNoteHome.jsx";
import { ENoteNewForm } from "./pages/eNoteNewForm.jsx";
import { ENoteViewForm } from "./pages/eNoteViewForm.jsx";
import Views from "./pages/eNoteViews.jsx";
import ATRViews from "./pages/ATRViews.jsx";
import { ENoteAtrworkflowform } from "./pages/eNoteAtrworkflowform.jsx";
import { AuthenticatedTemplate, UnauthenticatedTemplate, useIsAuthenticated } from "@azure/msal-react";
import { SearchNotes } from "./pages/searchnotes.jsx";
import { ENotePasscode } from "./pages/eNotePasscode.jsx"
import { useMsalAuthentication } from "@azure/msal-react";
import { InteractionType } from "@azure/msal-browser";
import { useMsal, useAccount } from '@azure/msal-react';
import PageNotFound from "./pages/404page.jsx";
import AdminDashBoardChartReports from "./pages/DashbaordReportforAdmin.jsx";
import { UploadSignature } from "./pages/uploadSignature.jsx";
import { ApproveSignature } from "./pages/approveSignature.jsx";  

// Function to get access token
export const getAccessToken = async (loginRequest, instance) => {
  try {
    const response = await instance.acquireTokenSilent(loginRequest);
    return response.accessToken;
  } catch (error) {
    console.error("Failed to get access token:", error);
    throw new Error("Failed to get access token");
  }
}

const TabContext = createContext();

export const useTabContext = () => useContext(TabContext);
function App() {
  const isAuthenticated = useIsAuthenticated();
  const { instance, accounts, inProgress } = useMsal();
  const account = useAccount(accounts[0] || {});
  const [activeTab, setActiveTab] = useState('My Notes'); // Default active tab
   const [activeButtons, setActiveButtons] = useState('My Pending Notes'); // Default active button
  const [passcodenaviagate, setPasscodeNaviage] = useState('New'); // Default passcode page redirect
  const [passcodeData, setPasscodeData] = useState(null);
  const [isFromUrl, setIsFromUrl] = useState(false);
   const [sessionStartTime, setSessionStartTime] = useState(null);
  const [sessionExpired, setSessionExpired] = useState(false);

  // Added for page redirect to my peding notes after any action
  const setTab = (tabName) => {
    setActiveTab(tabName);
  };
  // Added for page redirect to view page after creating passcode --> Kavya(23/07)
  const setPasscodeNavigate = (pageName) => {
    setPasscodeNaviage(pageName);
  }

   useEffect(() => {
    const rawTab = new URLSearchParams(window.location.search).get('tab');
    const tabParam = rawTab ? decodeURIComponent(rawTab.replace(/\+/g, " ")) : null;
    console.log("Detected tab param from URL:", tabParam);

    if (tabParam) {
      setActiveTab(tabParam);
      setActiveButtons(tabParam);
      setIsFromUrl(true);
    } else {
      setActiveTab('My Notes'); // Default active tab if no param
    }
  }, []);

  // Auto redirect to login
  useMsalAuthentication(InteractionType.Redirect);

  useEffect(() => {
    async function getTokenSilently() {
      const tokenRequest = {
        account,
        scopes: [process.env.REACT_APP_Scopes],
      };

      await instance.acquireTokenSilent(tokenRequest);
    }
    if (isAuthenticated && inProgress === 'none') {
      getTokenSilently();
    }
  }, [isAuthenticated]);

  useEffect(() => {
    if (isAuthenticated && inProgress === "none") {
      if (!sessionStartTime) {
        setSessionStartTime(Date.now());  // Start memory-based timer
      }
    }
  }, [isAuthenticated, inProgress]);

  useEffect(() => {
    if (!sessionStartTime) return;

    const thirteenHours = 13 * 60 * 60 * 1000; // 13 hours in ms

    const interval = setInterval(() => {
      const now = Date.now();
      if (now - sessionStartTime >= thirteenHours) {
        setSessionExpired(true);
      }
    }, 60000); // Check every minute

    return () => clearInterval(interval);
  }, [sessionStartTime]);
  
  return (
    <div className="App">
      <AuthenticatedTemplate>
        <TabContext.Provider value={{ activeTab, setTab, setPasscodeNavigate, passcodenaviagate, passcodeData, setPasscodeData,activeButtons, setActiveButtons, setIsFromUrl, isFromUrl }}>
        <Router>
          <Routes>
            <Route path="/" element={<ENoteHome />} />
            <Route path="/datagridpage" element={<ENoteHome />} />
            <Route path="/searchnotes" element={<SearchNotes />} />
            <Route path="/enoteform/:id" element={<ENoteNewForm />} />
            <Route path="/enoteviewform/:id" element={<ENoteViewForm />} />
            <Route path="/views/:id" element={<Views />} />
            <Route path="/atrviews/:id" element={<ATRViews />} />
            <Route path="/eNoteAtrworkflowform/:id" element={<ENoteAtrworkflowform />} />
            <Route  path="/dashBoardReports" element={<AdminDashBoardChartReports />}/>
             <Route path="/enotepasscode" element={<ENotePasscode />} /> 
             <Route path="/landingpage" element={<Landingpage />} />
              <Route path="/uploadSignature" element={<UploadSignature />} />
            <Route path="/approveSignature" element={<ApproveSignature />} />
            <Route path="*"  element={<PageNotFound />}/>
          </Routes>
        </Router>
        </TabContext.Provider>
      </AuthenticatedTemplate>
      <UnauthenticatedTemplate>
      <TabContext.Provider value={{ activeTab, setTab, setPasscodeNavigate, passcodenaviagate, passcodeData, setPasscodeData  }}>
        <Router>
          <Routes>
            <Route path="/" element={<Landingpage />} />
          </Routes>
        </Router>
        </TabContext.Provider>
      </UnauthenticatedTemplate>
       {sessionExpired && (
        <Dialog
          title={"Session Expired"}
          onClose={() => {
            setSessionExpired(false);
            instance.logoutRedirect();
          }}
        >
          <p>Your session has expired. Please login again to continue.</p>

          <DialogActionsBar>
             <Button
              className="notifyDailogOkBtn k-button k-button-md k-rounded-md k-button-solid k-button-solid-base"
              onClick={() => {
                setSessionExpired(false);
                instance.logoutRedirect(); // MSAL logout
              }}
            >
              <span>Login Again</span>
            </Button>
          </DialogActionsBar>
        </Dialog>
      )}
    </div>
  );
}

export default App;