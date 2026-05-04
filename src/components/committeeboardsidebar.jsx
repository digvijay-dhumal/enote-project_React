import React, { useState } from 'react';
import { ExpansionPanel, ExpansionPanelContent, } from "@progress/kendo-react-layout";
import { Link } from 'react-router-dom';
import { Button } from '@progress/kendo-react-buttons';
import toggle from '../../src/assets/toggle.svg'
import { Reveal } from "@progress/kendo-react-animation";
import { committeboarddatagridlistitems } from '../pages/datagridlistitems'
export const CommitteeBoardSidebar = () => {
    const [isFirstDivVisible, setFirstDivVisibility] = useState(false);
    const [isToggleClicked, setIsToggleClicked] = useState(false);
    const [expanded, setExpanded] = useState(false);
    const handleToggleHover = () => {
        setFirstDivVisibility(!isFirstDivVisible);
        setIsToggleClicked((prev) => !prev);
    };
    const getLinkForItem = (item) => {
        // Add your logic here to determine the link based on the item
        if (item === committeboarddatagridlistitems[0].items[0]) {
            return "/committeenewform";
        } else if (item === committeboarddatagridlistitems[0].items[1]) {
            return "/boardnewform"; // Add your logic for the second set of items
        } else {
            return ""; // Default link if no specific condition is met
        }
    };
    return (
        <div >
            <div style={{ height: '100%', position: 'fixed', zIndex: 2, width: '20%', background: '#f5f5f5', paddingTop: '20px', display: isFirstDivVisible ? 'block' : 'none' }}>
                {committeboarddatagridlistitems.map(item => (
                    <ExpansionPanel style={{ margin: '10px' }}
                        title={
                            <div style={{ display: 'flex' }}>
                                <div style={{ width: '18px', marginRight: '2px' }} dangerouslySetInnerHTML={{ __html: item.image }} />
                                {item.id}
                            </div>
                        }
                        expanded={expanded === item.id}
                        tabIndex={0}
                        onAction={event => {
                            setExpanded(event.expanded ? '' : item.id);
                        }}>

                        <Reveal>
                            {expanded === item.id && <ExpansionPanelContent>

                                <ul style={{ paddingLeft: '0px', marginBottom: '0px' }}>
                                    {item.items.map((listItem, index) => (
                                        <Link to={getLinkForItem(listItem)} style={{ textDecoration: 'none' }}>
                                            <li className="list-item" style={{ listStyleType: 'none' }} key={index}>{listItem}</li>
                                        </Link>
                                    ))}
                                </ul>

                            </ExpansionPanelContent>}
                        </Reveal>

                    </ExpansionPanel>
                ))}
            </div>
            <div
                style={{
                    paddingTop: '20px',
                    paddingLeft: isToggleClicked ? '20%' : '30px',
                    paddingRight: isToggleClicked ? 'auto' : 'auto',
                    position: 'absolute',
                    zIndex: 2,
                }}
            >
                {!isToggleClicked && (
                    <Button style={{ width: '60px' }} onClick={handleToggleHover}>Home</Button>
                )}
                {isToggleClicked && (
                    <img
                        src={toggle}
                        alt="Close Icon"
                        onClick={handleToggleHover}
                    />
                )}
            </div>
        </div>
    )
}