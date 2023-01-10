import React, { useState, useEffect, Fragment } from 'react';
import Helmet from 'react-helmet';
import axios from 'axios';
import './reports.css';


const Reports = () => {
    const [data, setData] = useState(null);
    const [loading, setLoading] = useState(true);
    const [test, setTest] = useState(null);
    useEffect(() => {
        const fetchData = async () => {
            const result = await axios('/api/dailyPost', {
                baseURL: 'http://localhost:5000'
            });
            console.log(`Response from server: ${JSON.stringify(result.data)}`);
            setData(result.data);
        };
        const fetchTestPy = async () => {
            const result = await axios('api/hellopy', {
                baseURL: 'http://localhost:5000'
            });
            console.log(`Response from server: ${JSON.stringify(result.data)}`);
            setTest(result.data);
        };
        (async () => {
            await fetchData();
            await fetchTestPy();
            setLoading(false);
        })();
    }, []);
    //Spit the values from fetchData by <br> and display them

    if (loading) {
        return (
            <div className='spinner-container'>
                <div className="spinner" />
            </div>
        )
    }

    return (
        <div className='page'>
            <Helmet>
                <style>
                    {`
                    body {
                        background - color: #FF0000;
                    background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='100%25' height='100%25' viewBox='0 0 1600 800'%3E%3Cg %3E%3Cpath fill='%23e60000' d='M486 705.8c-109.3-21.8-223.4-32.2-335.3-19.4C99.5 692.1 49 703 0 719.8V800h843.8c-115.9-33.2-230.8-68.1-347.6-92.2C492.8 707.1 489.4 706.5 486 705.8z'/%3E%3Cpath fill='%23cc0000' d='M1600 0H0v719.8c49-16.8 99.5-27.8 150.7-33.5c111.9-12.7 226-2.4 335.3 19.4c3.4 0.7 6.8 1.4 10.2 2c116.8 24 231.7 59 347.6 92.2H1600V0z'/%3E%3Cpath fill='%23b30000' d='M478.4 581c3.2 0.8 6.4 1.7 9.5 2.5c196.2 52.5 388.7 133.5 593.5 176.6c174.2 36.6 349.5 29.2 518.6-10.2V0H0v574.9c52.3-17.6 106.5-27.7 161.1-30.9C268.4 537.4 375.7 554.2 478.4 581z'/%3E%3Cpath fill='%23990000' d='M0 0v429.4c55.6-18.4 113.5-27.3 171.4-27.7c102.8-0.8 203.2 22.7 299.3 54.5c3 1 5.9 2 8.9 3c183.6 62 365.7 146.1 562.4 192.1c186.7 43.7 376.3 34.4 557.9-12.6V0H0z'/%3E%3Cpath fill='%23800000' d='M181.8 259.4c98.2 6 191.9 35.2 281.3 72.1c2.8 1.1 5.5 2.3 8.3 3.4c171 71.6 342.7 158.5 531.3 207.7c198.8 51.8 403.4 40.8 597.3-14.8V0H0v283.2C59 263.6 120.6 255.7 181.8 259.4z'/%3E%3Cpath fill='%23770d0d' d='M1600 0H0v136.3c62.3-20.9 127.7-27.5 192.2-19.2c93.6 12.1 180.5 47.7 263.3 89.6c2.6 1.3 5.1 2.6 7.7 3.9c158.4 81.1 319.7 170.9 500.3 223.2c210.5 61 430.8 49 636.6-16.6V0z'/%3E%3Cpath fill='%236d1b1b' d='M454.9 86.3C600.7 177 751.6 269.3 924.1 325c208.6 67.4 431.3 60.8 637.9-5.3c12.8-4.1 25.4-8.4 38.1-12.9V0H288.1c56 21.3 108.7 50.6 159.7 82C450.2 83.4 452.5 84.9 454.9 86.3z'/%3E%3Cpath fill='%23632a2a' d='M1600 0H498c118.1 85.8 243.5 164.5 386.8 216.2c191.8 69.2 400 74.7 595 21.1c40.8-11.2 81.1-25.2 120.3-41.7V0z'/%3E%3Cpath fill='%23573a3a' d='M1397.5 154.8c47.2-10.6 93.6-25.3 138.6-43.8c21.7-8.9 43-18.8 63.9-29.5V0H643.4c62.9 41.7 129.7 78.2 202.1 107.4C1020.4 178.1 1214.2 196.1 1397.5 154.8z'/%3E%3Cpath fill='%234B4B4B' d='M1315.3 72.4c75.3-12.6 148.9-37.1 216.8-72.4h-723C966.8 71 1144.7 101 1315.3 72.4z'/%3E%3C/g%3E%3C/svg%3E");
                    background-attachment: fixed;
                    background-size: cover;
                    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Roboto', 'Oxygen',
    'Ubuntu', 'Cantarell', 'Fira Sans', 'Droid Sans', 'Helvetica Neue',
    sans-serif;
                }
            `}
                </style>
            </Helmet>
            <div className="reports">
                <h1 className="report-title">Reports</h1>
                <div className="report">
                    <h1 className="report-subtitle">Daily Postcard Report</h1>
                    <table className='dailyPost-table'>
                        <tr>
                            <th className='table-header'>Postcard Sales</th>
                        </tr>
                        {data.data.split('\r\n').map((item, key) => {
                            if (item.includes(':')) {
                                return <Fragment key={key}><tr><td className="line-with-colon" >{item}</td></tr></Fragment>
                            } else {
                                return <Fragment key={key}><tr><td className="line-without-colon" >{item}</td></tr></Fragment>
                            }

                        })}
                    </table>

                </div>
                <div className="report">
                    <h1 className='report-title'>Test Py</h1>
                    <p className='report-data'>Data: {JSON.stringify(test)}</p>
                </div>
            </div>
        </div>
    );
};

export default Reports;