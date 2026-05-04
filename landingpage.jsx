import * as React from 'react';
import { Button } from "@progress/kendo-react-buttons";
import landingimage1 from '../../src/assets/landingimage1.png';
export const Landingpage = () => {
    return (
        <>
            <div class="section-1">
                <div >
                    <h1>Lorem Ipsum is simply dummy text</h1>
                    <p>Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.</p>
                    <Button>Get Started</Button>
                </div>
                <div>
                <img alt="IB image" src={landingimage1} />
                </div>
            </div>
            <div>Section 2</div>
        </>
    )
}