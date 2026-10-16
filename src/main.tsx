import React from "react";
import ReactDOM from "react-dom/client";
import { App as AntApp, ConfigProvider } from "antd";
import zhCN from "antd/locale/zh_CN";
import App from "./App";
import "./styles.css";

ReactDOM.createRoot(document.getElementById("root")!).render(
  <React.StrictMode>
    <ConfigProvider
      locale={zhCN}
      theme={{
        token: {
          colorPrimary: "#0687ad",
          colorInfo: "#0687ad",
          colorSuccess: "#159b8b",
          borderRadius: 9,
          fontFamily: 'Inter, "Segoe UI", "Microsoft YaHei", sans-serif',
          fontSize: 13,
          controlHeight: 36,
          colorText: "#233e50",
          colorBorder: "#dbe6ed",
        },
        components: {
          Button: { primaryShadow: "0 4px 12px rgba(0,133,174,.14)" },
          Drawer: { paddingLG: 28 },
        },
      }}
    >
      <AntApp>
        <App />
      </AntApp>
    </ConfigProvider>
  </React.StrictMode>,
);
