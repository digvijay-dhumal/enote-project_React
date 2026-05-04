import * as React from 'react';
import { ReorderContext } from '../components/tableDragRows';
import { useDraggable, useDroppable } from '@progress/kendo-react-common';
export const DraggableRow = props => {
  const [dropped, setDropped] = React.useState(false);
  const [dragged, setDragged] = React.useState(false);
  const [direction, setDirection] = React.useState(null);
  const [initial, setInitial] = React.useState({
    x: 0,
    y: 0
  });
  const {
    dragStart,
    reorder
  } = React.useContext(ReorderContext);
  const element = React.useRef(null);
  const handlePress = event => {
    setInitial({
      x: event.clientX - event.offsetX,
      y: event.clientY - event.offsetY
    });
  };
  const handleDragStart = event => {
    if (!event.originalEvent.target || !event.originalEvent.target.dataset.dragHandle) {
      return;
    }
    setDragged(true);
    dragStart(props.dataItem);
  };
  const handleDrag = event => {
    if (!element.current || !dragged) {
      return;
    }
    element.current.style.transform = `translateY(${event.clientY - initial.y + event.scrollY}px)`;
  };
  const handleDragEnd = () => {
    setDragged(false);
    setDropped(false);
    setInitial({
      x: 0,
      y: 0
    });
  };
  const handleRelease = () => {
    if (!element.current) {
      return;
    }
    element.current.style.transform = null;
  };
  const handleDragEnter = () => {
    setDropped(true);
    setDirection(null);
  };
  const handleDragOver = event => {
    if (!element.current) {
      return;
    }
    const rect = element.current.getBoundingClientRect();
    setDirection(rect.top + rect.height / 2 <= event.pageY ? 'after' : 'before');
  };
  const handleDragLeave = () => {
    setDropped(false);
    setDirection(null);
  };
  const handleDrop = () => {
    reorder(props.dataItem, direction);
    setDropped(false);
    setDirection(null);
  };
  useDraggable(element, {
    onPress: handlePress,
    onDragStart: handleDragStart,
    onDrag: handleDrag,
    onDragEnd: handleDragEnd,
    onRelease: handleRelease
  }, {
    autoScroll: dragged
  });
  useDroppable(element, {
    onDragEnter: handleDragEnter,
    onDragOver: handleDragOver,
    onDragLeave: handleDragLeave,
    onDrop: handleDrop
  });
  const rowIndex = props.dataIndex;
  const isEvenRow = rowIndex % 2 === 0;
  const rowHeight = 50; 
  return <React.Fragment>
        {dropped && direction === 'before' && (<tr style={{
      outlineStyle: 'solid',
      outlineWidth: 1,
      outlineColor: 'red'
    }} />)}
        <tr {...props.elementProps} ref={element} style={{
      backgroundColor: isEvenRow ? '#f5f5f5' : '#ffffff',
      height: `${rowHeight}px`,
      userSelect: 'none',
      pointerEvents: dragged ? 'none' : undefined,
      opacity: dragged ? '0.8' : undefined
    }} />
        {dropped && direction === 'after' && (<tr style={{
      outlineStyle: 'solid',
      outlineWidth: 1,
      outlineColor: 'red'
    }} />)}
      </React.Fragment>;
};