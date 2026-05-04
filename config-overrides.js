const {override} = require('customize-cra');
const cspHtmlWebpackPlugin = require("csp-html-webpack-plugin");
 
const cspConfigPolicy = {
    'default-src': "'self'",
    'base-uri': "'self'",
    'style-src': ["'self'"],
     'script-src': ["'self'", "blob:"],
    'img-src': ["'self'", "data:", "blob:"], 
    'connect-src': [
        process.env.REACT_APP_APIBase_URL,
        process.env.REACT_APP_eNote_URL,
        "https://login.microsoftonline.com",
        "blob:",
        "https://indianbank-webapp-aiteam-ceezd6dghdaheqe0.canadacentral-01.azurewebsites.net"
    ],
    'font-src': ["'self'"],
    'frame-src': ["'self'"],
    'media-src': ["'self'"]
};
 
function addCspHtmlWebpackPlugin(config) {
    // config.plugins.push(new cspHtmlWebpackPlugin(cspConfigPolicy));
    return config;
}
 
module.exports = {
    webpack: override(addCspHtmlWebpackPlugin),
};