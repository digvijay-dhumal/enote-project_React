import React, { useState } from 'react';
import { Button } from "@progress/kendo-react-buttons";
import Navbar from "../../src/components/navbar";
import Footer from "../components/footer";
import { Link } from 'react-router-dom';
import landingimage1 from '../../src/assets/homeBg.jpg';
import landingimage2 from '../../src/assets/landingimage2.png';
import landingimage3 from '../../src/assets/landingimage3.png';
import landingimage4 from '../../src/assets/landingimage4.png';
import landingimage5 from '../../src/assets/Image_For_e-Archieve.jpg';
import feedback from '../../src/assets/feedback.png';
import contact from '../../src/assets/contact.png';
import help from '../../src/assets/help.png'
import { Card, CardBody } from "@progress/kendo-react-layout";
import '../styles/landingpage.css';
import "bootstrap/dist/css/bootstrap.min.css";
import { useIsAuthenticated } from '@azure/msal-react';
import { IB_eNoteDomain_URL, IB_eCommitteeDomain_URL, IB_eDakDomain_URL } from '../config';


export const Landingpage = () => {
    const [centerSectionContent, setCenterSectionContent] = useState('IB Smart Office');

    const handleCardClick = () => {
      setCenterSectionContent('IB Smart Office - eDak');
    };
    const isAuthenticated = useIsAuthenticated();
    return (
        <>
            <Navbar header={centerSectionContent} />
            
            <>
                <div className="section-1">
                    <div className='container section-1a'>
                        <div className='section-1b'>
                            <h4>Smart Office</h4>
                            <p>An office space that anticipates your needs, optimizes your work flow and boosts your wellbeing. Presenting IB Smart Office, a tech-driven environment designed to enhance productivity, efficiency, minimize paper clutter and enhance employee satisfaction . It helps to faster document access, improved collaboration, and a more eco-friendly work environment.</p>
                        </div>
                        <div className='section-1c'>
                            <img alt="limg1" src={landingimage1} className='homeBg'/>
                        </div>
                        <div className='feedbackbuttons-container' >
                        {isAuthenticated ?(<button className="button">
                                <img className="icon" alt="Feedback" src={feedback} />  <span className="text">Feedback</span>
                            </button>):""}
                        </div>
                        <div className='contactbuttons-container' >
                        {isAuthenticated ?(<button className="button">
                                <img className="icon" alt="Feedback" src={contact} />  <span className="text">Contact Us</span>
                            </button>):""}
                        </div>
                    </div>
                </div>
                <div className="container">
                    <div className='row' style={{ marginTop: '30px' }}>
                        <div className='col-md-3'>
                        {isAuthenticated ? (<Link to={IB_eNoteDomain_URL} style={{textDecoration:'none'}}>
                            <Card className="custom-card" onClick={handleCardClick}>
                                <CardBody className="custom-cardimg" >
                                    <img alt="E-Note" src={landingimage2} />
                                </CardBody>
                                <Button className="full-width-button" >E-Note</Button>
                            </Card>
                            </Link>
                               ) : (
                                <Card onClick={handleCardClick} className={`custom-card ${isAuthenticated ? '' : 'faded-card'}`} >
                                <CardBody className="custom-cardimg" >
                                    <img alt="E-Note" src={landingimage2} />
                                </CardBody>
                                <Button className="full-width-button">E-Note</Button>
                            </Card>
                            )}
                        </div>
                        <div className='col-md-3'>
                        {isAuthenticated ? ( <Link to={IB_eCommitteeDomain_URL} style={{textDecoration:'none'}}>
                            <Card className="custom-card">
                                <CardBody className="custom-cardimg" >
                                    <img alt="Commitee/Board Notes" src={landingimage3} />
                                </CardBody>
                                <Button className="full-width-button">Commitee/Board Notes</Button>
                            </Card>
                            </Link>):(<Card className={`custom-card ${isAuthenticated ? '' : 'faded-card'}`} >
                                <CardBody className="custom-cardimg" >
                                    <img alt="Commitee/Board Notes" src={landingimage3} />
                                </CardBody>
                                <Button className="full-width-button">Commitee/Board Notes</Button>
                            </Card>)}
                        </div>
                       <div className='col-md-3'>
                        {isAuthenticated ? (<Link to={IB_eDakDomain_URL} style={{textDecoration:'none'}}>
                            <Card className="custom-card" onClick={handleCardClick}>
                                <CardBody className="custom-cardimg" >
                                    <img alt="eDak" src={landingimage4} />
                                </CardBody>
                                <Button className="full-width-button" >E-Dak</Button>
                           
                            </Card>
                            </Link>
                               ) : (
                                <Card onClick={handleCardClick} className={`custom-card ${isAuthenticated ? '' : 'faded-card'}`} >
                                <CardBody className="custom-cardimg" >
                                    <img alt="eDak" src={landingimage4} />
                                </CardBody>
                                <Button className="full-width-button">E-Dak</Button>
                            </Card>
                            )}
                        </div>
                        <div className='col-md-3'>
                        {isAuthenticated?(
                            <Link to="https://indianbanko365.sharepoint.com/sites/smartoffice" style={{textDecoration:'none'}}>
                                <Card className="custom-card">
                                    <CardBody className="custom-cardimg">
                                        <img alt="E-Archival" src={landingimage5} className='img_archival' />
                                    </CardBody>
                                    <Button className="full-width-button">E-Archival</Button>
                                </Card>
                            </Link>):( <Link to="https://indianbanko365.sharepoint.com/sites/smartoffice" style={{textDecoration:'none'}}><Card className={`custom-card ${isAuthenticated ? '' : 'faded-card'}`}>
                                <CardBody className="custom-cardimg" >
                                    <img alt="E-Archival" src={landingimage5}  className='img_archival'/>
                                </CardBody>
                                <Button className=" ebuttons full-width-button">E-Archival</Button>
                            </Card>
                         </Link>)}
                        </div>
                    </div>
                </div>
                <div className="container section-3">
                    <Button >Help <img alt="help" src={help} /></Button>
                </div>
            </>
            <div className="pgFooterContainer">
                <Footer />
            </div>
        </>
    )
}