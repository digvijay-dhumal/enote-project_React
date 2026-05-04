import * as React from "react";
import {
    Chart,
    ChartLegend,
    ChartSeries,
    ChartSeriesItem,
    ChartTitle,
    ChartTooltip
} from "@progress/kendo-react-charts";
import "hammerjs";
import "bootstrap/dist/css/bootstrap.min.css";

const tooltipRender = props => {
    if (props.point) {
      const {
        category,
      } = props.point;
      return <span>{String(category)}</span>;
    }
  };
export const ChartContainer = (props) => {

    const [state, setState] = React.useState({})
    React.useEffect(() => {
        setState(props.apiOutput)
    }, [props]);
   
    return (
        <div className="custom-committeeNote-charts-container">
            <div className="row">
                {state ?
                    <>
                        <div className="col-lg-6 col-md-12 col-sm-12 custom-pie-chart-seperateline ">
                            <div className="enote-chartContainer">
                                <Chart>
                                    <ChartTitle text="Status" />
                                    <ChartLegend position="right">
                                    </ChartLegend>
                                    <ChartSeries>
                                        <ChartSeriesItem
                                            type="pie"
                                            data={state.lstNoteChart}
                                            field="count"
                                            name="status"
                                            categoryField="status"

                                            tooltip={{
                                                visible: true,
                                            }}
                                        />

                                    </ChartSeries>
                                    <ChartTooltip opacity={1} visible={true} render={tooltipRender}/>
                                </Chart>
                            </div>
                        </div>
                        <div className="col-lg-6 col-md-12 col-sm-12">
                            <div className="enote-chartContainer">
                                <Chart
                                >
                                    <ChartTitle text="Nature of Notes Created" />
                                    <ChartLegend
                                        position="right"
                                    />
                                    <ChartSeries>
                                        <ChartSeriesItem
                                            type="donut"
                                            data={state.lstNatureofNoteStatusCount}
                                            field="count"
                                            categoryField="status"
                                            tooltip={{
                                                visible: true,
                                            }}
                                        />
                                    </ChartSeries>
                                    <ChartTooltip opacity={1} visible={true} render={tooltipRender}/>
                                </Chart>
                            </div>
                        </div>
                    </>
                    :
                    <div>
                        <span>
                            No data is availble.
                        </span>
                    </div>}
            </div>
        </div>
    )
}