import { useMsal } from '@azure/msal-react';
import { loginRequest } from "../config.js";
import { Button } from '@progress/kendo-react-buttons';
export const SignInButton = () => {
  const { instance, accounts } = useMsal();

  const handleLogin = (loginType) => {
  if (loginType === "redirect") {
      instance.loginRedirect(loginRequest).catch((e) => {
        console.log(accounts[0].name);
      });
    }
  };
  return (
    <div>
      {accounts.length === 0 && 
      (
         <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', height: '100vh' }}>
        <p style={{marginBottom:'22px'}} >Welcome! Please click on Sign In to proceed with application</p>
        <Button style={{height:'40px',width:'50px',background:'#F0F8FF'}} onClick={() => handleLogin("redirect")}>Sign In</Button>
 
      </div>
      )
       }
    </div>
  );
};
