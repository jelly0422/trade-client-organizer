import React from "react";
import ReactDOM from "react-dom/client";
import { App as AntApp, ConfigProvider } from "antd";
import zhCN from "antd/locale/zh_CN";
import App from "./App";
import "./styles.css";
import "./glass-theme.css";

ReactDOM.createRoot(document.getElementById("root")!).render(
  <React.StrictMode>
    <ConfigProvider
      locale={zhCN}
      theme={{
        token: {
          colorPrimary: "#087da5",
          colorInfo: "#087da5",
          colorSuccess: "#159b8b",
          borderRadius: 10,
          fontFamily: 'Inter, "Segoe UI", "Microsoft YaHei", sans-serif',
          fontSize: 13,
          controlHeight: 36,
          colorText: "#12394d",
          colorTextSecondary: "#516c7d",
          colorBorder: "#c9dfe7",
        },
        components: {
          Button: { primaryShadow: "0 5px 12px rgba(9,127,160,.11)" },
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
