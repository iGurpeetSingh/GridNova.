import { Navigate, Route, Routes } from "react-router-dom";
import Dashboard from "./pages/Dashboard";
import WorkbookEditor from "./pages/WorkbookEditor";

export default function App() {
  return (
    <Routes>
      <Route path="/" element={<Dashboard />} />
      <Route path="/workbook/:id" element={<WorkbookEditor />} />
      <Route path="*" element={<Navigate to="/" replace />} />
    </Routes>
  );
}