import React from 'react';
import { BrowserRouter, Routes, Route } from 'react-router-dom';
import Navbar from './navbar';
import Home from './Home';
import About from './About';
import APITest from './APITest';
import Reports from './reports'
const App = () => {
  return (
    <BrowserRouter>
    <Navbar />
    <Routes>
      <Route path="/" element={<Home/>}/>
        <Route path="/about" element={<About/>} />
        <Route exact path="/api-test" element={<APITest/>} />
        <Route exact path="/reports" element={<Reports/>} />
      {/* </Route> */}
      </Routes>
    </BrowserRouter>
  );
};

export default App;