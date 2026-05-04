import * as React from 'react';
import { reorderIcon } from '@progress/kendo-svg-icons';
import { SvgIcon } from '@progress/kendo-react-common';
export const DragHandleCell = props => {
  return <td {...props} style={{
    touchAction: 'none'
  }}>
      <span style={{
      cursor: 'move'
    }} data-drag-handle={true}>
        <SvgIcon style={{
        pointerEvents: 'none'
      }} icon={reorderIcon} />
      </span>
    </td>;
};