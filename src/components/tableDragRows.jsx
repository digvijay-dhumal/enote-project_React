import React, { useContext, useState, useRef, useEffect } from 'react';
import { Grid, GridColumn as Column } from '@progress/kendo-react-grid';
import { SvgIcon } from '@progress/kendo-react-common';
import { reorderIcon } from '@progress/kendo-svg-icons';
import { Button } from "@progress/kendo-react-buttons";
import '../styles/responsiveDesign.css'

// Context to handle the dragging logic
const ReorderContext = React.createContext({
  reorder: (dataItem) => {},
  dragStart: (dataItem) => {},
});

const DragCell = (props) => {
  const currentContext = useContext(ReorderContext);
  const touchMoveRef = useRef(null);

  // Handle drag start for mouse events
  const handleDragStart = (e) => {
    currentContext.dragStart(props.dataItem);
    e.dataTransfer.setData('dragging', '');
  };

  // Handle touch start for mobile devices
  const handleTouchStart = () => {
    currentContext.dragStart(props.dataItem);

    // Add the touchmove listener dynamically when dragging starts
    touchMoveRef.current = handleTouchMove;
    window.addEventListener('touchmove', touchMoveRef.current, { passive: false });
  };

  // Handle touch move for mobile devices
  const handleTouchMove = (e) => {
    const touch = e.touches[0];
    const target = document.elementFromPoint(touch.clientX, touch.clientY);
    if (target && target.closest('tr')) {
      const row = target.closest('tr');
      const index = row.rowIndex - 1; // Subtract 1 for the header row
      currentContext.reorder(props.dataItem, index);
    }
    e.preventDefault();
  };

  // Clean up the touchmove listener when touch ends
  const handleTouchEnd = () => {
    window.removeEventListener('touchmove', touchMoveRef.current);
  };

  useEffect(() => {
    // Attach touchend listener to cleanup after drag is complete
    window.addEventListener('touchend', handleTouchEnd);
    return () => {
      window.removeEventListener('touchend', handleTouchEnd);
      window.removeEventListener('touchmove', touchMoveRef.current);
    };
  }, []);

  return (
    <td
      onDragOver={(e) => e.preventDefault()}
      onDrop={(e) => {
        currentContext.reorder(props.dataItem);
        e.preventDefault();
      }}
      onTouchStart={handleTouchStart}
    >
      <span
        draggable={true}
        style={{ cursor: 'move' }}
        onDragStart={handleDragStart}
      >
        <SvgIcon icon={reorderIcon} />
      </span>
    </td>
  );
};

// Main component for the grid with drag-and-drop functionality
export const TableDraggableRows = (props) => {
  const{reviewerCol,srCol,designationCol}= props;
  const [gridData, setGridData] = useState(props.data);
  const [gridHeight, setGridHeight] = useState('auto');
  const [activeItem, setActiveItem] = useState(null);

  React.useEffect(() => {
    setGridData(props.data)   
  }, [props])

  const updatedDate = (current) => {
    const updatedRow = gridData.filter((obj, index) => index !== current);

    props.onDelate(gridData.find((obj,ind)=>ind===current));
    setGridData(updatedRow);
  }

  const reorder = (dataItem, targetIndex = null) => {
    if (activeItem === dataItem) {
      return;
    }
  
    // Create a copy of the grid data to avoid mutating state directly
    let reorderedData = [...gridData];
    
    // Find the previous index of the active item
    const prevIndex = reorderedData.findIndex(p => p === activeItem);
    
    // If targetIndex is not provided, find it based on dataItem
    if (targetIndex === null) {
      targetIndex = reorderedData.findIndex(p => p === dataItem);
    }
  
    if (prevIndex > -1 && targetIndex > -1) {
      // Remove the active item from its original position
      reorderedData.splice(prevIndex, 1);
      
      // Insert the active item at the new target position
      reorderedData.splice(targetIndex, 0, activeItem);
  
      // Update the approverOrder for each item based on the new order
      const updatedData = reorderedData.map((obj, ind) => ({
        ...obj,
        approverOrder: ind + 1,
      }));
  
      // Update the grid data with the reordered items
      setGridData(updatedData);
  
      // Notify the parent component about the order change if needed
      if (props.onOrderChange) {
        props.onOrderChange(updatedData);
      }
    }
  };
  

  // Function to start the drag operation
  const dragStart = (dataItem) => {
    setActiveItem(dataItem);
  };

  return (
    <div style={{overflowX:"auto",overflowY: "auto !important",maxHeight:"100vh" }}>
      <ReorderContext.Provider value={{ reorder: reorder, dragStart: dragStart }}>
        <Grid
          style={{ height: gridHeight }}
          data={gridData}
          dataItemKey={"ProductID"}
          rowClassName={(index) => (index % 2 === 0 ? "even-row" : "odd-row")}
        >
          <Column title="" width="40px" cell={DragCell} />
          <Column  title="S.No" width="100" cell={(props) => {
            return <td>{props.dataIndex + 1}</td>;
         }} />
        {/* Change - Adding SR No - 05/04 - RK  */}
          {/*  Change - Adding Designation- 05/04 - RK  */}
        <Column field="approverEmailName" title={reviewerCol} className='mobile_responsive_dragTable'/>
        <Column field="srNo" width="150" title={srCol}/>
      
        <Column field="designation" title={designationCol} className='mobile_responsive_dragTable'/>
        <Column width="120" cell={props => {
          return (<td>
            <Button
              onClick={() => {  updatedDate(props?.dataIndex)}}    
            >
              <span className="k-icon-sm k-icon k-font-icon k-i-delete k-i-trash cursor allIconsforPrimary-btn"></span>
              Delete
            </Button>
          </td>)
        }} title="Action" />

        </Grid>
      </ReorderContext.Provider>
    </div>
  );
};