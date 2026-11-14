import type { PluginClientContext } from "@getpaseo/plugin/client";
import { ExcelFilePanel } from "./client/panel";

export default function contribute(client: PluginClientContext) {
  // One panel; quick review and the formatted view switch from its toolbar.
  const removePanel = client.addFilePanel({
    id: "excel",
    title: "Excel 보기",
    icon: "Sheet",
    files: { extensions: ["xlsx", "xlsm", "xltx", "xltm", "xls", "xlt"] },
    Component: ExcelFilePanel,
  });
  return () => {
    removePanel();
  };
}
