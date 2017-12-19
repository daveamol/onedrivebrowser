import * as React from 'react';
import './StatusBar.css';

class StatusBar extends React.Component {
  render() {
    return (
      <div className="status-bar">
        <div className="ms-BrandIcon--icon96 ms-BrandIcon--onedrive App-icon" />
        <span className="ms-font-xl ms-fontWeight-semibold">One Drive Browser</span>
      </div>
    );
  }
}

export default StatusBar;