import './storageShim.js';
import React from 'react';
import ReactDOM from 'react-dom/client';
import './index.css';
import MoneyTracker from './App.jsx';

ReactDOM.createRoot(document.getElementById('root')).render(
  <React.StrictMode>
    <MoneyTracker />
  </React.StrictMode>
);
