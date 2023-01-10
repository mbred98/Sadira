import React from 'react';
import { Link } from 'react-router-dom';
import './navbar.css';
import logo from './decroixintl.png'

const Navbar = () => {
  return (
    <nav className="navbar">
      <Link to="/">
        <img src={logo} alt="Logo" className='navbar-logo'/>
      </Link>
      <div>
        <Link to="/" className="navbar-link">
          Home
        </Link>
        <Link to="/about" className="navbar-link">
          About
        </Link>
        <Link to="/api-test" className="navbar-link">
          API Test
        </Link>
        <Link to="/reports" className="navbar-link">
          Reports
        </Link>
      </div>
    </nav>
  );
};

export default Navbar;