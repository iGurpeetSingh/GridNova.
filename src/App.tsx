import { lazy, Suspense } from "react";
import { Navigate, Route, Routes } from "react-router-dom";

const Dashboard = lazy(() => import("./pages/Dashboard"));
const WorkbookEditor = lazy(() => import("./pages/WorkbookEditor"));

function RouteLoading() {
  return <div className="app-shell grid min-h-screen place-items-center text-sm font-semibold text-[#718078]">Opening GridNova…</div>;
}
export default function App() {
  return <Suspense fallback={<RouteLoading />}>
    <Routes>
      <Route path="/" element={<Dashboard />} />
      <Route path="/workbook/:id" element={<WorkbookEditor />} />
      <Route path="*" element={<Navigate to="/" replace />} />
    </Routes>
  </Suspense>;
}