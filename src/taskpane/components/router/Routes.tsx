import React, { useEffect, useState } from "react";
import { HashRouter as Router, Routes, Route } from "react-router-dom";
import Home from "../layout/Main/Home";
import Getstart from "../layout/Main/GetStart";

const RouterApp: React.FC = () => {
    const [itemChange, setItemChanged] = useState<string>("");
    const [lastItem, setLastItem] = useState<string>("");

    useEffect(() => {
        const initializeOffice = async () => {
            await Office.onReady(); // Ensure Office.js is ready before running any logic

            const itemChanged = () => {
                const item = Office.context.mailbox?.item;
                if (item?.itemId && item.itemId !== lastItem) {
                    setLastItem(item.itemId);
                    setItemChanged(item.itemId);
                }
            };

            // Add event listener
            Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, itemChanged);

            // Cleanup function to remove event listener
            return () => {
                Office.context.mailbox.removeHandlerAsync(Office.EventType.ItemChanged, itemChanged);
            };
        };

        initializeOffice();
    }, [lastItem]);

    return (
        <Router>
            <Routes>
                <Route path="/" element={<Getstart />} />
                <Route path="/Home" element={<Home selectdItemFromAdrees={itemChange} />} />
            </Routes>
        </Router>
    );
};

export default RouterApp;
