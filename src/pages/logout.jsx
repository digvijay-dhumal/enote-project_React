import { useMsal } from '@azure/msal-react';
import { IB_Domain_URL } from '../config.js';


export const Logout = (props) => {
  const { instance } = useMsal();

  instance.logoutRedirect({
    postLogoutRedirectUri: IB_Domain_URL
  });
};