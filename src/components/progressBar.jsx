import { ProgressBar } from "@progress/kendo-react-progressbars";
import "../styles/footer.css"

export const LoadingOverlay = ({ value }) => (
    <div className="overlay">
        <div className="overlay-content">
            <ProgressBar
                value={value}
                labelVisible={true}
                label={() => `${value}%`}
                className="custom-progress-bar"
            />
        </div>
    </div>
);
