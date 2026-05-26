import React, { useState, useEffect } from "react";
import {
  AreaChart,
  Area,
  BarChart,
  Bar,
  PieChart,
  Pie,
  Cell,
  ComposedChart,
  Line,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
  ResponsiveContainer,
} from "recharts";
import {
  DollarSign,
  ShieldAlert,
  Users,
  MapPin,
  MessageSquare,
  Clock,
  CheckCircle2,
  AlertTriangle,
  Activity,
  Zap,
  ListChecks,
  Download,
  Lock,
  Plus,
  X,
} from "lucide-react";

// Dynamically import xlsx for Excel handling
let XLSX: any = null;
if (typeof window !== "undefined") {
  import("xlsx").then((mod) => {
    XLSX = mod;
  });
}

// --- DATA ---
const defaultData = {
  spendData: [
    {
      month: "Apr-25",
      reactive: 66292,
      ppm: 133864,
      quoted: 41414,
      materials: 23441,
      total: 241572,
    },
    {
      month: "May-25",
      reactive: 40347,
      ppm: 134571,
      quoted: 61850,
      materials: 53441,
      total: 290210,
    },
    {
      month: "Jun-25",
      reactive: 31234,
      ppm: 136274,
      quoted: 123932,
      materials: 71242,
      total: 226408,
    },
    {
      month: "Jul-25",
      reactive: 47625,
      ppm: 139087,
      quoted: 132205,
      materials: 17254,
      total: 336172,
    },
    {
      month: "Aug-25",
      reactive: 60943,
      ppm: 141182,
      quoted: 79965,
      materials: 29513,
      total: 311604,
    },
    {
      month: "Sep-25",
      reactive: 59646,
      ppm: 142371,
      quoted: 94982,
      materials: 24752,
      total: 321753,
    },
    {
      month: "Oct-25",
      reactive: 74532,
      ppm: 147919,
      quoted: 83290,
      materials: 19684,
      total: 325426,
    },
    {
      month: "Nov-25",
      reactive: 68668,
      ppm: 147919,
      quoted: 119042,
      materials: 39381,
      total: 375011,
    },
    {
      month: "Dec-25",
      reactive: 57114,
      ppm: 147919,
      quoted: 884286,
      materials: 36437,
      total: 1125757,
    },
    {
      month: "Jan-26",
      reactive: 25457,
      ppm: 147919,
      quoted: 91383,
      materials: 21900,
      total: 294624,
    },
    {
      month: "Feb-26",
      reactive: 86253,
      ppm: 147919,
      quoted: 86398,
      materials: 14280,
      total: 336433,
    },
    {
      month: "Mar-26",
      reactive: 61280,
      ppm: 155459,
      quoted: 80717,
      materials: 19966,
      total: 318477,
    },
    {
      month: "Apr-26",
      reactive: 42256,
      ppm: 145441,
      quoted: 67073,
      materials: 9798,
      total: 267569,
    },
  ],
  intruder2026: [
    {
      month: "Jan-26",
      totalSpend: 5814,
      jobsInvoiced: 32,
      visits: 28,
      remoteResets: 4,
    },
    {
      month: "Feb-26",
      totalSpend: 8983,
      jobsInvoiced: 49,
      visits: 47,
      remoteResets: 2,
    },
    {
      month: "Mar-26",
      totalSpend: 7264,
      jobsInvoiced: 50,
      visits: 25,
      remoteResets: 25,
    },
    {
      month: "Apr-26",
      totalSpend: 6858,
      jobsInvoiced: 64,
      visits: 25,
      remoteResets: 39,
    },
  ],
  slaPerformance: [
    { priority: "P1 (Emergency)", logged: 72, completed: 72, percent: 100 },
    { priority: "P2 (Urgent)", logged: 2, completed: 0, percent: 0 },
    { priority: "P3 (Important)", logged: 12, completed: 7, percent: 58.3 },
    { priority: "P4 (Routine)", logged: 2, completed: 2, percent: 100 },
    { priority: "P5 (Low)", logged: 5, completed: 5, percent: 100 },
  ],
  ppmBreakdown: [
    { name: "UKN Engineers", due: 649, completed: 647, percent: 99.69 },
    { name: "Subcontractors", due: 23, completed: 20, percent: 86.96 },
  ],
  missedTasksLog: [
    { team: "UKN Engineers", task: "Henfield: 1 Monthly Flushing overlooked." },
    {
      team: "UKN Engineers",
      task: "Northallerton: 6 Monthly Roller Shutter service.",
    },
    {
      team: "Subcontractors",
      task: "Belfast: 1 Monthly Fire Alarm (Site declined access).",
    },
    { team: "Subcontractors", task: "Newtownards: 1 Monthly Fire Alarm." },
  ],
  topReactiveBranches2026: [
    { site: "GSF Burgess Hill 603 (CLOSED)", spend: 9768.95, jobs: 9 },
    { site: "GSF Plymouth South 425", spend: 4923.41, jobs: 23 },
    { site: "GSF Birmingham - Castle Vale", spend: 4374.37, jobs: 13 },
    { site: "GSF Woodford Green 783", spend: 3329.53, jobs: 5 },
    { site: "GSF Saltash 434", spend: 3309.88, jobs: 5 },
  ],
  topIntruderBranches: [
    { site: "GSF Grimsby 270", spend: 1014.0, jobs: 2 },
    { site: "GSF Hemel Hempstead 104", spend: 985.8, jobs: 7 },
    { site: "GSF Staples Corner 781", spend: 972.0, jobs: 6 },
    { site: "GSF Merthyr 402", spend: 757.5, jobs: 3 },
    { site: "GSF Tottenham 780", spend: 720.0, jobs: 4 },
  ],
  cxFeedbackLog: [
    {
      site: "Hemel Hempstead",
      issue:
        "Fire alarm tests need to be recorded in file after tests. The quality of service does vary depending on who you get.",
      resolution:
        "The GSF file should be completed by GSF following weekly tests. We have a completed job sheet for the monthly test which is shared with our risk assessor.",
    },
    {
      site: "Castle Vale",
      issue:
        "Often we have no idea what they have turned up for, sometimes you can see 2/3 a week and often seem to stumble over each others work.",
      resolution:
        "Attended this site 23 times in 2025. We do combine where possible, for example, gas drop and boiler, roller shutter and POU were combined.",
    },
    {
      site: "Leeds West",
      issue:
        "Sometimes several UKN employees attend branch within a short timescale instead of consolidating into one visit.",
      resolution:
        "We do combine as much as possible. The only time we have attended twice in one week is when monthly PPMs were completed and then a separate reactive job was raised.",
    },
    {
      site: "Bridgend",
      issue:
        "We can have checks done at the end of a month and then again the following week. Faulty equipment takes too long to be fixed.",
      resolution:
        "We have never completed the monthly visit this close together. The closest has been 20 days apart. Alarm resets escalated to UKSI.",
    },
    {
      site: "Hayes",
      issue:
        "Alarm code resets take far too long. Had to wait weeks to get one reset recently.",
      resolution:
        "Most jobs resolved same day by phone or attended promptly (e.g. logged 13/10 8am, attended 14/10 12pm). 1 awaiting UKSI paperwork.",
    },
  ],
};

// --- UTILS ---
const formatMoney = (val: number) =>
  new Intl.NumberFormat("en-GB", {
    style: "currency",
    currency: "GBP",
    maximumFractionDigits: 0,
  }).format(val);

// --- CUSTOM INTERACTIVE TOOLTIP ---
const CustomTooltip = ({
  active = false,
  payload = [],
  label = "",
}: {
  active?: boolean;
  payload?: Array<{ name: string; value: number; color: string }>;
  label?: string;
}) => {
  if (active && payload && payload.length) {
    return (
      <div className="bg-white/95 backdrop-blur-sm p-4 rounded-xl shadow-2xl border border-slate-100 transform transition-all z-50">
        <p className="font-bold text-slate-800 mb-2 border-b border-slate-100 pb-1">
          {label}
        </p>
        {payload.map(
          (
            entry: { name: string; value: number; color: string },
            index: number
          ) => {
            const isNum =
              entry.name.includes("percent") ||
              entry.name.includes("Resets") ||
              entry.name.includes("Visits") ||
              entry.name.includes("Tasks") ||
              entry.name.includes("Due") ||
              entry.name.includes("Completed");
            return (
              <div
                key={index}
                className="flex items-center space-x-3 text-sm my-1"
              >
                <div
                  className="w-3 h-3 rounded-full shadow-inner"
                  style={{ backgroundColor: entry.color }}
                ></div>
                <span className="text-slate-600 font-medium">
                  {entry.name}:
                </span>
                <span className="font-bold text-slate-900">
                  {isNum ? entry.value : formatMoney(entry.value)}
                </span>
              </div>
            );
          }
        )}
      </div>
    );
  }
  return null;
};

// --- VIBRANT CARD COMPONENT ---
interface CardProps {
  title: string;
  value: string | number;
  subtext: string;
  icon: React.ComponentType<React.SVGProps<SVGSVGElement>>;
  gradientFrom: string;
  gradientTo: string;
  iconColor: string;
  onClick?: () => void;
}

const Card = ({
  title,
  value,
  subtext,
  icon: Icon,
  gradientFrom,
  gradientTo,
  iconColor,
  onClick = undefined,
}: CardProps) => {
  const containerClass =
    "bg-gradient-to-br rounded-3xl shadow-lg border border-white/40 p-6 transition-all duration-300 hover:-translate-y-2 hover:shadow-2xl cursor-pointer relative overflow-hidden group " +
    gradientFrom +
    " " +
    gradientTo;
  const iconClass =
    "w-8 h-8 transition-transform duration-300 group-hover:scale-110 " +
    iconColor;

  return (
    <div className={containerClass} onClick={onClick}>
      <div className="absolute -right-8 -top-8 w-32 h-32 bg-white opacity-20 rounded-full blur-2xl group-hover:scale-150 transition-transform duration-700"></div>
      <div className="flex justify-between items-start relative z-10">
        <div>
          <h3 className="text-slate-700/80 text-xs font-bold uppercase tracking-widest mb-2">
            {title}
          </h3>
          <div className="text-4xl font-black text-slate-800 mb-3 drop-shadow-sm tracking-tight">
            {value}
          </div>
          <p className="text-slate-700 text-xs font-bold flex items-center bg-white/50 backdrop-blur-sm inline-block px-3 py-1.5 rounded-lg shadow-sm">
            {subtext}
          </p>
        </div>
        <div className="p-4 rounded-2xl bg-white/80 backdrop-blur-md shadow-sm transform group-hover:rotate-12 transition-transform duration-300">
          <Icon className={iconClass} />
        </div>
      </div>
    </div>
  );
};

// --- MAIN APP ---
export default function App() {
  const [activeTab, setActiveTab] = useState("executive");

  // Initialize data as state for dynamic updates
  const [data, setData] = useState(() => {
    if (typeof window !== "undefined") {
      const savedData = window.localStorage.getItem("gsfDashboardData");
      if (savedData) {
        try {
          const parsedData = JSON.parse(savedData);
          return { ...defaultData, ...parsedData };
        } catch (e) {
          console.error("Error parsing saved data:", e);
          return defaultData;
        }
      }
    }
    return defaultData;
  });

  useEffect(() => {
    if (typeof window !== "undefined") {
      window.localStorage.setItem("gsfDashboardData", JSON.stringify(data));
    }
  }, [data]);

  // Auth & Edit State
  const [isAuthenticated, setIsAuthenticated] = useState(false);
  const [showAuthModal, setShowAuthModal] = useState(false);
  const [password, setPassword] = useState("");
  const [showAddModal, setShowAddModal] = useState(false);
  const [selectedTable, setSelectedTable] =
    useState<keyof typeof defaultData>("spendData");
  const [formData, setFormData] = useState<any>({});
  const [editingIndex, setEditingIndex] = useState<number | null>(null);

  // Destructure state values cleanly
  const {
    spendData,
    intruder2026,
    slaPerformance,
    ppmBreakdown,
    missedTasksLog,
    topReactiveBranches2026,
    topIntruderBranches,
    cxFeedbackLog,
  } = data;

  // Group missed tasks by team
  const groupedMissedTasks = (missedTasksLog || []).reduce((acc, task) => {
    if (!acc[task.team]) {
      acc[task.team] = [];
    }
    acc[task.team].push(task);
    return acc;
  }, {} as Record<string, typeof missedTasksLog>);

  // Derived state for latest month's breakdown
  const latestSpendData =
    spendData && spendData.length > 0 ? spendData[spendData.length - 1] : null;
  const latestSpendBreakdown = latestSpendData
    ? [
        { name: "PPM", value: latestSpendData.ppm, color: "#0ea5e9" },
        { name: "Quoted", value: latestSpendData.quoted, color: "#8b5cf6" },
        { name: "Reactive", value: latestSpendData.reactive, color: "#f97316" },
        {
          name: "Materials",
          value: latestSpendData.materials,
          color: "#ec4899",
        },
      ]
    : [];

  const downloadTemplate = () => {
    if (!XLSX) {
      alert(
        "Excel export library is still loading. Please try again in a moment."
      );
      return;
    }
    try {
      const workbook = XLSX.utils.book_new();

      // Add each data sheet
      Object.keys(defaultData).forEach((key) => {
        XLSX.utils.book_append_sheet(
          workbook,
          XLSX.utils.json_to_sheet((data as any)[key]),
          key
            .replace(/([A-Z])/g, " $1")
            .replace(/^./, (str) => str.toUpperCase()) // Prettify sheet name
        );
      });

      // Generate filename with timestamp
      const now = new Date();
      const timestamp = now.toISOString().split("T")[0];
      const filename = `GSF_BusinessReview_Template_${timestamp}.xlsx`;

      // Write file
      XLSX.writeFile(workbook, filename);
    } catch (error) {
      console.error("Error downloading template:", error);
      alert("Error downloading template. Check console for details.");
    }
  };

  const handleLogin = () => {
    if (password === "admin123") {
      setIsAuthenticated(true);
      setShowAuthModal(false);
      setPassword("");
    } else {
      alert("Incorrect password!");
    }
  };

  const handleCancelEdit = () => {
    setEditingIndex(null);
    setFormData({});
  };

  const handleEditClick = (index: number) => {
    setEditingIndex(index);
    setFormData(data[selectedTable][index]);
  };

  const handleDeleteClick = (index: number) => {
    if (
      window.confirm(
        `Are you sure you want to delete this entry?\n\n${Object.values(
          data[selectedTable][index]
        ).join(" | ")}`
      )
    ) {
      setData((prev) => ({
        ...prev,
        [selectedTable]: (prev[selectedTable] || []).filter(
          (_, i) => i !== index
        ),
      }));
      if (editingIndex === index) {
        handleCancelEdit();
      }
    }
  };

  const handleSaveData = () => {
    const parsedData = { ...formData };
    Object.keys(parsedData).forEach((key) => {
      const value = parsedData[key];
      if (
        typeof value === "string" &&
        value.trim() !== "" &&
        !isNaN(Number(value))
      ) {
        parsedData[key] = Number(value);
      }
    });

    if (selectedTable === "spendData") {
      const { reactive = 0, ppm = 0, quoted = 0, materials = 0 } = parsedData;
      parsedData.total =
        Number(reactive) + Number(ppm) + Number(quoted) + Number(materials);
    }

    if (editingIndex !== null) {
      // Update
      const tableData = [...(data[selectedTable] || [])];
      tableData[editingIndex] = parsedData;
      setData((prev) => ({ ...prev, [selectedTable]: tableData }));
      alert("Data updated successfully!");
    } else {
      // Add
      setData((prev) => ({
        ...prev,
        [selectedTable]: [...(prev[selectedTable] || []), parsedData],
      }));
      alert("Data added successfully!");
    }

    setShowAddModal(false);
    handleCancelEdit();
  };

  const tabs = [
    { id: "executive", label: "Executive Summary", icon: Activity },
    { id: "financials", label: "Financial Spend", icon: DollarSign },
    { id: "ppm", label: "PPM Breakdown", icon: ListChecks },
    { id: "operations", label: "Operations & SLAs", icon: Zap },
    { id: "cx", label: "Customer Experience", icon: Users },
  ];

  return (
    <div className="min-h-screen bg-[#f1f5f9] font-sans text-slate-900 pb-12">
      {/* VIBRANT HEADER */}
      <header className="bg-white/80 backdrop-blur-lg border-b border-slate-200 sticky top-0 z-30 shadow-sm">
        <div className="max-w-7xl mx-auto px-6 py-4 flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div className="flex items-center space-x-4">
            <div className="w-14 h-14 bg-white rounded-xl flex items-center justify-center shadow-md border border-slate-100 transform transition hover:scale-105 cursor-pointer overflow-hidden p-1.5">
              <img
                src="images.png"
                alt="GSF Logo"
                className="w-full h-full object-contain"
                onError={(e) => {
                  const target = e.target as HTMLImageElement;
                  target.onerror = null;
                  target.src =
                    "https://placehold.co/100x100/ef4444/ffffff?text=GSF";
                }}
              />
            </div>
            <div>
              <h1 className="text-2xl font-black bg-clip-text text-transparent bg-gradient-to-r from-slate-800 to-slate-500 tracking-tight">
                GSF Business Review
              </h1>
              <p className="text-xs font-bold text-red-600 tracking-widest uppercase mt-0.5">
                May 2026 • UK National Ltd
              </p>
            </div>
          </div>

          <div className="flex flex-col xl:flex-row items-center space-y-4 xl:space-y-0 xl:space-x-4">
            {/* ACTIONS */}
            <div className="flex space-x-2">
              <button
                onClick={() => {
                  if (
                    window.confirm(
                      "Are you sure you want to reset all data to the default template? This cannot be undone."
                    )
                  ) {
                    setData(defaultData);
                    alert("Data has been reset to the default template.");
                  }
                }}
                className="flex items-center space-x-2 px-4 py-2 bg-yellow-50 text-yellow-600 rounded-full hover:bg-yellow-100 cursor-pointer font-bold text-xs transition-colors border border-yellow-200 shadow-sm"
              >
                <AlertTriangle className="w-4 h-4" />
                <span className="hidden sm:inline">Reset Data</span>
              </button>

              <button
                onClick={downloadTemplate}
                className="flex items-center space-x-2 px-4 py-2 bg-emerald-50 text-emerald-600 rounded-full hover:bg-emerald-100 cursor-pointer font-bold text-xs transition-colors border border-emerald-200 shadow-sm"
              >
                <Download className="w-4 h-4" />
                <span className="hidden sm:inline">Get Template</span>
              </button>

              {!isAuthenticated ? (
                <button
                  onClick={() => setShowAuthModal(true)}
                  className="flex items-center space-x-2 px-4 py-2 bg-slate-50 text-slate-600 rounded-full hover:bg-slate-100 cursor-pointer font-bold text-xs transition-colors border border-slate-200 shadow-sm"
                >
                  <Lock className="w-4 h-4" />
                  <span className="hidden sm:inline">Admin Login</span>
                </button>
              ) : (
                <div className="flex space-x-2">
                  <button
                    onClick={() => setShowAddModal(true)}
                    className="flex items-center space-x-2 px-4 py-2 bg-blue-50 text-blue-600 rounded-full hover:bg-blue-100 cursor-pointer font-bold text-xs transition-colors border border-blue-200 shadow-sm"
                  >
                    <Plus className="w-4 h-4" />
                    <span className="hidden sm:inline">Add Data</span>
                  </button>
                  <button
                    onClick={() => setIsAuthenticated(false)}
                    className="flex items-center space-x-2 px-4 py-2 bg-rose-50 text-rose-600 rounded-full hover:bg-rose-100 cursor-pointer font-bold text-xs transition-colors border border-rose-200 shadow-sm"
                  >
                    <span className="hidden sm:inline">Log Out</span>
                  </button>
                </div>
              )}
            </div>

            <div className="flex space-x-1 bg-slate-100/80 p-1.5 rounded-full shadow-inner overflow-x-auto border border-slate-200/50 max-w-full">
              {tabs.map((tab) => {
                const isActive = activeTab === tab.id;
                const btnBase =
                  "flex items-center space-x-2 px-4 py-2.5 text-sm font-bold rounded-full transition-all duration-300 whitespace-nowrap ";
                const activeClass =
                  btnBase +
                  "bg-white text-blue-600 shadow-md transform scale-100";
                const inactiveClass =
                  btnBase +
                  "text-slate-500 hover:text-slate-800 hover:bg-slate-200/50";

                return (
                  <button
                    key={tab.id}
                    onClick={() => setActiveTab(tab.id)}
                    className={isActive ? activeClass : inactiveClass}
                  >
                    <tab.icon
                      className={"w-4 h-4 " + (isActive ? "animate-pulse" : "")}
                    />
                    <span>{tab.label}</span>
                  </button>
                );
              })}
            </div>
          </div>
        </div>
      </header>

      {/* CONTENT WITH ANIMATED APPEARANCE */}
      <main className="px-6 max-w-7xl mx-auto mt-8 animate-fade-in-up">
        {/* TAB 1: EXECUTIVE */}
        {activeTab === "executive" && (
          <div className="space-y-8 animate-[fadeIn_0.5s_ease-out]">
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
              <Card
                title={`Total Spend (${
                  latestSpendData ? latestSpendData.month.split("-")[0] : "N/A"
                })`}
                value={
                  latestSpendData ? formatMoney(latestSpendData.total) : "£0"
                }
                subtext="▼ Down from Last Mo."
                icon={DollarSign}
                gradientFrom="from-blue-50"
                gradientTo="to-cyan-100"
                iconColor="text-blue-600"
                onClick={() => setActiveTab("financials")}
              />
              <Card
                title="PPM Completion"
                value={ppmBreakdown[0]?.percent + "%"}
                subtext="UKN Engineers"
                icon={CheckCircle2}
                gradientFrom="from-emerald-50"
                gradientTo="to-teal-100"
                iconColor="text-emerald-600"
                onClick={() => setActiveTab("ppm")}
              />
              <Card
                title="P1 Compliance"
                value={slaPerformance[0]?.percent + "%"}
                subtext="Emergency SLA"
                icon={Clock}
                gradientFrom="from-indigo-50"
                gradientTo="to-purple-100"
                iconColor="text-indigo-600"
                onClick={() => setActiveTab("operations")}
              />
              <Card
                title="CX Score"
                value="99%"
                subtext="Latest Surveys"
                icon={Users}
                gradientFrom="from-amber-50"
                gradientTo="to-orange-100"
                iconColor="text-amber-600"
                onClick={() => setActiveTab("cx")}
              />
            </div>

            <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
              <div className="lg:col-span-2 bg-white rounded-3xl shadow-xl shadow-slate-200/50 border border-slate-100 p-8 hover:shadow-2xl hover:shadow-slate-200/50 transition-shadow duration-300">
                <div className="flex justify-between items-center mb-8">
                  <div>
                    <h3 className="text-2xl font-black text-slate-800 tracking-tight">
                      Total Spend Trend
                    </h3>
                    <p className="text-sm font-medium text-slate-500 mt-1">
                      13-Month historical view
                    </p>
                  </div>
                  <span className="bg-blue-50 text-blue-600 border border-blue-100 px-4 py-1.5 rounded-full text-xs font-bold uppercase tracking-wider shadow-sm">
                    All Services
                  </span>
                </div>
                <div className="h-80">
                  <ResponsiveContainer width="100%" height="100%">
                    <AreaChart
                      data={spendData}
                      margin={{ top: 10, right: 10, left: 0, bottom: 0 }}
                    >
                      <defs>
                        <linearGradient
                          id="colorTotalActive"
                          x1="0"
                          y1="0"
                          x2="0"
                          y2="1"
                        >
                          <stop
                            offset="5%"
                            stopColor="#3b82f6"
                            stopOpacity={0.5}
                          />
                          <stop
                            offset="95%"
                            stopColor="#8b5cf6"
                            stopOpacity={0.0}
                          />
                        </linearGradient>
                      </defs>
                      <CartesianGrid
                        strokeDasharray="4 4"
                        vertical={false}
                        stroke="#e2e8f0"
                      />
                      <XAxis
                        dataKey="month"
                        axisLine={false}
                        tickLine={false}
                        tick={{
                          fill: "#64748b",
                          fontSize: 11,
                          fontWeight: 600,
                        }}
                        dy={10}
                        interval={0}
                      />
                      <YAxis
                        tickFormatter={(val) => "£" + val / 1000 + "k"}
                        axisLine={false}
                        tickLine={false}
                        tick={{
                          fill: "#64748b",
                          fontSize: 13,
                          fontWeight: 600,
                        }}
                        dx={-10}
                      />
                      <Tooltip
                        content={<CustomTooltip />}
                        cursor={{
                          stroke: "#94a3b8",
                          strokeWidth: 2,
                          strokeDasharray: "4 4",
                        }}
                      />
                      <Area
                        type="monotone"
                        dataKey="total"
                        stroke="#3b82f6"
                        strokeWidth={4}
                        fill="url(#colorTotalActive)"
                        name="Total Spend"
                        activeDot={{ r: 8, strokeWidth: 3, stroke: "#fff" }}
                      />
                    </AreaChart>
                  </ResponsiveContainer>
                </div>
              </div>

              <div className="bg-white rounded-3xl shadow-xl shadow-slate-200/50 border border-slate-100 p-8 hover:shadow-2xl hover:shadow-slate-200/50 transition-shadow duration-300">
                <h3 className="text-2xl font-black text-slate-800 mb-1 tracking-tight">
                  Latest Breakdown
                </h3>
                <p className="text-sm font-medium text-slate-500 mb-6">
                  Spend allocation
                </p>
                <div className="h-[320px]">
                  <ResponsiveContainer width="100%" height="100%">
                    <PieChart>
                      <Pie
                        data={latestSpendBreakdown}
                        cx="50%"
                        cy="50%"
                        innerRadius={80}
                        outerRadius={110}
                        paddingAngle={8}
                        dataKey="value"
                        stroke="none"
                        cornerRadius={6}
                      >
                        {latestSpendBreakdown.map((entry, index) => (
                          <Cell
                            key={index}
                            fill={entry.color || "#cbd5e1"}
                            className="hover:opacity-80 transition-opacity duration-200 cursor-pointer outline-none"
                          />
                        ))}
                      </Pie>
                      <Tooltip
                        content={
                          <CustomTooltip
                            active={undefined}
                            payload={undefined}
                            label={undefined}
                          />
                        }
                      />
                      <Legend
                        verticalAlign="bottom"
                        iconType="circle"
                        wrapperStyle={{
                          fontWeight: 600,
                          fontSize: "14px",
                          paddingTop: "20px",
                        }}
                      />
                    </PieChart>
                  </ResponsiveContainer>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* TAB 2: FINANCIALS */}
        {activeTab === "financials" && (
          <div className="space-y-8 animate-[fadeIn_0.5s_ease-out]">
            <div className="bg-white rounded-3xl shadow-xl shadow-slate-200/50 border border-slate-100 p-10">
              <div className="flex flex-col md:flex-row justify-between items-start md:items-center mb-10 gap-6">
                <div>
                  <h3 className="text-3xl font-black text-slate-800 tracking-tight">
                    Spend Composition
                  </h3>
                  <p className="text-slate-500 font-medium mt-2 text-sm">
                    Detailed 13-month breakdown of category expenditure.
                  </p>
                </div>
                <div className="flex flex-wrap gap-4 bg-slate-50 p-2 rounded-2xl border border-slate-100">
                  <span className="flex items-center px-3 py-1 bg-white rounded-xl shadow-sm text-sm font-bold text-slate-700">
                    <div className="w-3 h-3 rounded-full bg-[#0ea5e9] mr-2 shadow-inner"></div>
                    PPM
                  </span>
                  <span className="flex items-center px-3 py-1 bg-white rounded-xl shadow-sm text-sm font-bold text-slate-700">
                    <div className="w-3 h-3 rounded-full bg-[#f97316] mr-2 shadow-inner"></div>
                    Reactive
                  </span>
                  <span className="flex items-center px-3 py-1 bg-white rounded-xl shadow-sm text-sm font-bold text-slate-700">
                    <div className="w-3 h-3 rounded-full bg-[#8b5cf6] mr-2 shadow-inner"></div>
                    Quoted
                  </span>
                  <span className="flex items-center px-3 py-1 bg-white rounded-xl shadow-sm text-sm font-bold text-slate-700">
                    <div className="w-3 h-3 rounded-full bg-[#ec4899] mr-2 shadow-inner"></div>
                    Materials
                  </span>
                </div>
              </div>

              <div className="h-[450px]">
                <ResponsiveContainer width="100%" height="100%">
                  <AreaChart
                    data={spendData}
                    margin={{ top: 10, right: 10, left: 10, bottom: 10 }}
                  >
                    <defs>
                      <linearGradient id="colorPpm" x1="0" y1="0" x2="0" y2="1">
                        <stop
                          offset="5%"
                          stopColor="#0ea5e9"
                          stopOpacity={0.9}
                        />
                        <stop
                          offset="95%"
                          stopColor="#0ea5e9"
                          stopOpacity={0.3}
                        />
                      </linearGradient>
                      <linearGradient
                        id="colorReactive"
                        x1="0"
                        y1="0"
                        x2="0"
                        y2="1"
                      >
                        <stop
                          offset="5%"
                          stopColor="#f97316"
                          stopOpacity={0.9}
                        />
                        <stop
                          offset="95%"
                          stopColor="#f97316"
                          stopOpacity={0.3}
                        />
                      </linearGradient>
                      <linearGradient
                        id="colorQuoted"
                        x1="0"
                        y1="0"
                        x2="0"
                        y2="1"
                      >
                        <stop
                          offset="5%"
                          stopColor="#8b5cf6"
                          stopOpacity={0.9}
                        />
                        <stop
                          offset="95%"
                          stopColor="#8b5cf6"
                          stopOpacity={0.3}
                        />
                      </linearGradient>
                      <linearGradient
                        id="colorMaterials"
                        x1="0"
                        y1="0"
                        x2="0"
                        y2="1"
                      >
                        <stop
                          offset="5%"
                          stopColor="#ec4899"
                          stopOpacity={0.9}
                        />
                        <stop
                          offset="95%"
                          stopColor="#ec4899"
                          stopOpacity={0.3}
                        />
                      </linearGradient>
                    </defs>
                    <CartesianGrid
                      strokeDasharray="4 4"
                      vertical={false}
                      stroke="#e2e8f0"
                    />
                    <XAxis
                      dataKey="month"
                      tick={{ fill: "#64748b", fontWeight: 600, fontSize: 11 }}
                      axisLine={false}
                      tickLine={false}
                      dy={15}
                      interval={0}
                    />
                    <YAxis
                      tickFormatter={(val) => "£" + val / 1000 + "k"}
                      tick={{ fill: "#64748b", fontWeight: 600, fontSize: 13 }}
                      axisLine={false}
                      tickLine={false}
                      dx={-10}
                    />
                    <Tooltip
                      content={
                        <CustomTooltip
                          active={undefined}
                          payload={undefined}
                          label={undefined}
                        />
                      }
                      cursor={{ fill: "rgba(241, 245, 249, 0.5)" }}
                    />
                    <Area
                      type="monotone"
                      dataKey="ppm"
                      stackId="1"
                      stroke="#0284c7"
                      strokeWidth={2}
                      fill="url(#colorPpm)"
                      name="PPM"
                      activeDot={{ r: 6, strokeWidth: 0 }}
                    />
                    <Area
                      type="monotone"
                      dataKey="reactive"
                      stackId="1"
                      stroke="#ea580c"
                      strokeWidth={2}
                      fill="url(#colorReactive)"
                      name="Reactive"
                      activeDot={{ r: 6, strokeWidth: 0 }}
                    />
                    <Area
                      type="monotone"
                      dataKey="quoted"
                      stackId="1"
                      stroke="#7c3aed"
                      strokeWidth={2}
                      fill="url(#colorQuoted)"
                      name="Quoted"
                      activeDot={{ r: 6, strokeWidth: 0 }}
                    />
                    <Area
                      type="monotone"
                      dataKey="materials"
                      stackId="1"
                      stroke="#db2777"
                      strokeWidth={2}
                      fill="url(#colorMaterials)"
                      name="Materials"
                      activeDot={{ r: 6, strokeWidth: 0 }}
                    />
                  </AreaChart>
                </ResponsiveContainer>
              </div>
            </div>

            <div className="bg-white rounded-3xl shadow-xl shadow-slate-200/50 border border-slate-100 p-10 overflow-hidden relative">
              <div className="absolute top-0 right-0 w-64 h-64 bg-orange-50 rounded-full blur-3xl -mr-20 -mt-20 z-0"></div>
              <div className="relative z-10">
                <div className="flex items-center space-x-3 mb-6">
                  <div className="p-3 bg-orange-100 rounded-xl">
                    <DollarSign className="w-6 h-6 text-orange-600" />
                  </div>
                  <h3 className="text-2xl font-black text-slate-800 tracking-tight">
                    Top 5 Reactive Spend Branches (YTD)
                  </h3>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-5 gap-4">
                  {topReactiveBranches2026.map((branch, idx) => (
                    <div
                      key={idx}
                      className="bg-slate-50 border border-slate-200 rounded-2xl p-5 hover:border-orange-300 hover:shadow-md transition-all"
                    >
                      <div className="flex items-center space-x-2 mb-3">
                        <span className="flex items-center justify-center w-6 h-6 rounded-full bg-orange-200 text-orange-700 text-xs font-bold">
                          {idx + 1}
                        </span>
                        <MapPin className="w-4 h-4 text-slate-400" />
                      </div>
                      <p className="text-sm font-bold text-slate-800 mb-1 leading-tight min-h-[40px]">
                        {branch.site}
                      </p>
                      <div className="mt-4 pt-4 border-t border-slate-200">
                        <p className="text-2xl font-black text-orange-600 tracking-tight">
                          {formatMoney(branch.spend)}
                        </p>
                        <p className="text-xs font-semibold text-slate-500 mt-1 uppercase">
                          {branch.jobs} Jobs Invoiced
                        </p>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          </div>
        )}

        {/* TAB 3: PPM BREAKDOWN */}
        {activeTab === "ppm" && (
          <div className="space-y-8 animate-[fadeIn_0.5s_ease-out]">
            <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
              <div className="bg-gradient-to-br from-emerald-500 to-teal-600 p-8 rounded-3xl shadow-lg border border-emerald-400 text-white transform hover:-translate-y-1 transition-all">
                <p className="text-sm font-bold text-emerald-100 uppercase tracking-widest mb-2">
                  Overall Completion
                </p>
                <p className="text-5xl font-black mb-3">99.26%</p>
                <div className="bg-black/20 backdrop-blur-sm rounded-xl py-2 px-4 inline-block">
                  <p className="text-sm font-medium">
                    667 / 672 Tasks Completed
                  </p>
                </div>
              </div>
              <div className="bg-gradient-to-br from-blue-500 to-indigo-600 p-8 rounded-3xl shadow-lg border border-blue-400 text-white transform hover:-translate-y-1 transition-all">
                <p className="text-sm font-bold text-blue-100 uppercase tracking-widest mb-2">
                  {ppmBreakdown[0]?.name}
                </p>
                <p className="text-5xl font-black mb-3">
                  {ppmBreakdown[0]?.percent}%
                </p>
                <div className="bg-black/20 backdrop-blur-sm rounded-xl py-2 px-4 inline-block">
                  <p className="text-sm font-medium">
                    {ppmBreakdown[0]?.completed} / {ppmBreakdown[0]?.due} Tasks
                  </p>
                </div>
              </div>
              <div className="bg-gradient-to-br from-amber-500 to-orange-500 p-8 rounded-3xl shadow-lg border border-amber-400 text-white transform hover:-translate-y-1 transition-all">
                <p className="text-sm font-bold text-amber-100 uppercase tracking-widest mb-2">
                  {ppmBreakdown[1]?.name}
                </p>
                <p className="text-5xl font-black mb-3">
                  {ppmBreakdown[1]?.percent}%
                </p>
                <div className="bg-black/20 backdrop-blur-sm rounded-xl py-2 px-4 inline-block">
                  <p className="text-sm font-medium">
                    {ppmBreakdown[1]?.completed} / {ppmBreakdown[1]?.due} Tasks
                  </p>
                </div>
              </div>
            </div>

            <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
              <div className="bg-white p-8 rounded-3xl shadow-xl shadow-slate-200/50 border border-slate-100">
                <h3 className="text-2xl font-black text-slate-800 mb-2 tracking-tight">
                  Task Completion vs. Due
                </h3>
                <p className="text-sm font-medium text-slate-500 mb-8">
                  Performance metrics
                </p>
                <div className="h-80">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={ppmBreakdown} margin={{ left: -20 }}>
                      <CartesianGrid
                        strokeDasharray="4 4"
                        vertical={false}
                        stroke="#e2e8f0"
                      />
                      <XAxis
                        dataKey="name"
                        tick={{
                          fill: "#64748b",
                          fontWeight: 600,
                          fontSize: 13,
                        }}
                        axisLine={false}
                        tickLine={false}
                        dy={10}
                      />
                      <YAxis
                        axisLine={false}
                        tickLine={false}
                        tick={{ fill: "#64748b", fontWeight: 600 }}
                      />
                      <Tooltip
                        content={
                          <CustomTooltip
                            active={undefined}
                            payload={undefined}
                            label={undefined}
                          />
                        }
                        cursor={{ fill: "#f8fafc" }}
                      />
                      <Legend
                        wrapperStyle={{ paddingTop: "20px", fontWeight: 600 }}
                      />
                      <Bar
                        dataKey="due"
                        name="Tasks Due"
                        fill="#cbd5e1"
                        radius={[6, 6, 0, 0]}
                        barSize={60}
                      />
                      <Bar
                        dataKey="completed"
                        name="Completed"
                        fill="#3b82f6"
                        radius={[6, 6, 0, 0]}
                        barSize={60}
                      />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>

              <div className="bg-white p-8 rounded-3xl shadow-xl shadow-slate-200/50 border border-slate-100 relative overflow-hidden">
                <div className="absolute right-0 top-0 w-32 h-32 bg-rose-50 rounded-bl-full -z-10"></div>
                <div className="flex items-center space-x-3 mb-8 border-b border-slate-100 pb-4">
                  <div className="p-3 bg-rose-100 rounded-xl">
                    <AlertTriangle className="w-6 h-6 text-rose-600" />
                  </div>
                  <h3 className="text-2xl font-black text-slate-800 tracking-tight">
                    Missed Tasks Log
                  </h3>
                </div>

                <div className="space-y-6">
                  {Object.entries(groupedMissedTasks).map(([team, tasks]) => {
                    const isUkn = team.includes("UKN");
                    const barColor = isUkn ? "bg-blue-500" : "bg-amber-500";
                    const textColor = isUkn
                      ? "text-blue-700"
                      : "text-amber-700";
                    const bulletColor = isUkn
                      ? "text-blue-500"
                      : "text-amber-500";

                    return (
                      <div
                        key={team}
                        className="p-6 bg-slate-50 rounded-2xl border border-slate-200 shadow-sm"
                      >
                        <div className="flex items-center space-x-2 mb-3">
                          <div
                            className={`w-2 h-6 ${barColor} rounded-full`}
                          ></div>
                          <p
                            className={`text-sm font-black ${textColor} uppercase tracking-wider`}
                          >
                            {team} ({tasks.length} Missed)
                          </p>
                        </div>
                        <ul className="text-sm space-y-3 text-slate-700 font-medium">
                          {tasks.map((task, taskIdx) => {
                            const parts = task.task.split(":");
                            const location = parts[0];
                            const description = parts.slice(1).join(":").trim();
                            return (
                              <li key={taskIdx} className="flex items-start">
                                <span className={`${bulletColor} mr-2 mt-0.5`}>
                                  •
                                </span>
                                <span>
                                  <strong>{location}:</strong> {description}
                                </span>
                              </li>
                            );
                          })}
                        </ul>
                      </div>
                    );
                  })}
                </div>
              </div>
            </div>
          </div>
        )}

        {/* TAB 4: OPERATIONS & SLAs */}
        {activeTab === "operations" && (
          <div className="space-y-8 animate-[fadeIn_0.5s_ease-out]">
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
              {/* Intruder Spend & Visits Composed Chart */}
              <div className="bg-white p-8 rounded-3xl shadow-xl shadow-slate-200/50 border border-slate-100">
                <h3 className="text-2xl font-black text-slate-800 tracking-tight mb-2">
                  Intruder Callouts & Spend
                </h3>
                <p className="text-sm font-medium text-slate-500 mb-8">
                  YTD Financial & Visit Analysis
                </p>
                <div className="h-80">
                  <ResponsiveContainer width="100%" height="100%">
                    <ComposedChart
                      data={intruder2026}
                      margin={{ left: -20, right: 10 }}
                    >
                      <CartesianGrid
                        strokeDasharray="4 4"
                        vertical={false}
                        stroke="#e2e8f0"
                      />
                      <XAxis
                        dataKey="month"
                        tick={{
                          fill: "#64748b",
                          fontWeight: 600,
                          fontSize: 11,
                        }}
                        axisLine={false}
                        tickLine={false}
                        dy={10}
                        interval={0}
                      />
                      <YAxis
                        yAxisId="left"
                        tickFormatter={(v) => "£" + v / 1000 + "k"}
                        tick={{ fill: "#64748b", fontWeight: 600 }}
                        axisLine={false}
                        tickLine={false}
                        dx={-10}
                      />
                      <YAxis
                        yAxisId="right"
                        orientation="right"
                        tick={{ fill: "#64748b", fontWeight: 600 }}
                        axisLine={false}
                        tickLine={false}
                        dx={10}
                      />
                      <Tooltip
                        content={
                          <CustomTooltip
                            active={undefined}
                            payload={undefined}
                            label={undefined}
                          />
                        }
                        cursor={{ fill: "#f8fafc" }}
                      />
                      <Legend
                        iconType="circle"
                        wrapperStyle={{
                          fontWeight: 600,
                          paddingTop: "20px",
                          fontSize: "13px",
                        }}
                      />
                      <Bar
                        yAxisId="right"
                        dataKey="visits"
                        stackId="a"
                        name="Physical Visits"
                        fill="#f43f5e"
                        radius={[0, 0, 6, 6]}
                        barSize={45}
                      />
                      <Bar
                        yAxisId="right"
                        dataKey="remoteResets"
                        stackId="a"
                        name="Remote Resets"
                        fill="#0ea5e9"
                        radius={[6, 6, 0, 0]}
                        barSize={45}
                      />
                      <Line
                        yAxisId="left"
                        type="monotone"
                        dataKey="totalSpend"
                        name="Intruder Spend"
                        stroke="#10b981"
                        strokeWidth={5}
                        dot={{ r: 6, strokeWidth: 3, fill: "#fff" }}
                        activeDot={{ r: 8 }}
                      />
                    </ComposedChart>
                  </ResponsiveContainer>
                </div>
              </div>

              {/* SLA Compliance */}
              <div className="bg-white rounded-3xl shadow-xl shadow-slate-200/50 border border-slate-100 p-8 hover:shadow-2xl transition-shadow duration-300">
                <div className="flex justify-between items-center mb-8">
                  <div>
                    <h3 className="text-2xl font-black text-slate-800 tracking-tight">
                      SLA Compliance
                    </h3>
                    <p className="text-sm font-medium text-slate-500 mt-1">
                      Reactive response metrics
                    </p>
                  </div>
                  <div className="p-3 bg-slate-50 rounded-2xl border border-slate-100">
                    <ShieldAlert className="w-6 h-6 text-slate-400" />
                  </div>
                </div>

                <div className="space-y-7 mt-4">
                  {slaPerformance.map((item, idx) => {
                    const barColor =
                      item.percent >= 90
                        ? "#10b981"
                        : item.percent === 0
                        ? "#ef4444"
                        : "#f97316";
                    const barWidth = Math.max(item.percent, 2) + "%";

                    return (
                      <div key={idx} className="group">
                        <div className="flex justify-between items-end mb-2">
                          <span className="text-sm font-bold text-slate-700 group-hover:text-blue-600 transition-colors">
                            {item.priority}
                          </span>
                          <span
                            className="text-base font-black"
                            style={{ color: barColor }}
                          >
                            {item.percent}%{" "}
                            <span className="text-xs font-bold text-slate-400 ml-1">
                              ({item.completed}/{item.logged})
                            </span>
                          </span>
                        </div>
                        <div className="w-full bg-slate-100 rounded-full h-4 overflow-hidden shadow-inner border border-slate-200/50">
                          <div
                            className="h-full rounded-full transition-all duration-1000 ease-out relative overflow-hidden"
                            style={{
                              width: barWidth,
                              backgroundColor: barColor,
                            }}
                          >
                            <div className="absolute inset-0 bg-white/20 w-full h-full animate-[pulse_2s_ease-in-out_infinite]"></div>
                          </div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            </div>

            {/* Top Intruder Branches */}
            <div className="bg-white rounded-3xl shadow-xl shadow-slate-200/50 border border-slate-100 p-10 overflow-hidden relative">
              <div className="absolute top-0 right-0 w-64 h-64 bg-rose-50 rounded-full blur-3xl -mr-20 -mt-20 z-0"></div>
              <div className="relative z-10">
                <div className="flex items-center space-x-3 mb-6">
                  <div className="p-3 bg-rose-100 rounded-xl">
                    <ShieldAlert className="w-6 h-6 text-rose-600" />
                  </div>
                  <div>
                    <h3 className="text-2xl font-black text-slate-800 tracking-tight">
                      Top 5 Intruder Spend Branches
                    </h3>
                    <p className="text-sm font-medium text-slate-500 mt-1">
                      Last 3 Months Overview
                    </p>
                  </div>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-5 gap-4">
                  {topIntruderBranches.map((branch, idx) => (
                    <div
                      key={idx}
                      className="bg-slate-50 border border-slate-200 rounded-2xl p-5 hover:border-rose-300 hover:shadow-md transition-all"
                    >
                      <div className="flex items-center space-x-2 mb-3">
                        <span className="flex items-center justify-center w-6 h-6 rounded-full bg-rose-200 text-rose-700 text-xs font-bold">
                          {idx + 1}
                        </span>
                        <MapPin className="w-4 h-4 text-slate-400" />
                      </div>
                      <p className="text-sm font-bold text-slate-800 mb-1 leading-tight min-h-[40px]">
                        {branch.site}
                      </p>
                      <div className="mt-4 pt-4 border-t border-slate-200">
                        <p className="text-2xl font-black text-rose-600 tracking-tight">
                          {formatMoney(branch.spend)}
                        </p>
                        <p className="text-xs font-semibold text-slate-500 mt-1 uppercase">
                          {branch.jobs} Callouts
                        </p>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          </div>
        )}

        {/* TAB 5: CUSTOMER EXPERIENCE */}
        {activeTab === "cx" && (
          <div className="space-y-8 animate-[fadeIn_0.5s_ease-out]">
            <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
              <div className="bg-white rounded-3xl shadow-xl shadow-slate-200/50 border border-slate-100 p-8 flex flex-col items-center text-center group hover:-translate-y-2 transition-all duration-300">
                <div className="w-28 h-28 rounded-full border-8 border-amber-400 flex items-center justify-center mb-6 bg-gradient-to-br from-amber-50 to-orange-100 shadow-inner group-hover:scale-110 transition-transform duration-300">
                  <span className="text-4xl font-black text-amber-600 drop-shadow-sm">
                    99%
                  </span>
                </div>
                <h3 className="text-xl font-black text-slate-800 mt-2">
                  Professionalism
                </h3>
              </div>

              <div className="bg-white rounded-3xl shadow-xl shadow-slate-200/50 border border-slate-100 p-8 flex flex-col items-center text-center group hover:-translate-y-2 transition-all duration-300">
                <div className="w-28 h-28 rounded-full border-8 border-blue-400 flex items-center justify-center mb-6 bg-gradient-to-br from-blue-50 to-cyan-100 shadow-inner group-hover:scale-110 transition-transform duration-300">
                  <span className="text-4xl font-black text-blue-600 drop-shadow-sm">
                    95%
                  </span>
                </div>
                <h3 className="text-xl font-black text-slate-800 mt-2">
                  Uniform & PPE
                </h3>
              </div>

              <div className="bg-white rounded-3xl shadow-xl shadow-slate-200/50 border border-slate-100 p-8 flex flex-col items-center text-center group hover:-translate-y-2 transition-all duration-300">
                <div className="w-28 h-28 rounded-full border-8 border-emerald-400 flex items-center justify-center mb-6 bg-gradient-to-br from-emerald-50 to-teal-100 shadow-inner group-hover:scale-110 transition-transform duration-300">
                  <span className="text-4xl font-black text-emerald-600 drop-shadow-sm">
                    94%
                  </span>
                </div>
                <h3 className="text-xl font-black text-slate-800 mt-2">
                  Site Cleanliness
                </h3>
              </div>
            </div>

            <div className="bg-white rounded-3xl shadow-xl shadow-slate-200/50 border border-slate-100 p-10 relative overflow-hidden">
              <div className="mb-8">
                <h3 className="text-3xl font-black text-slate-800 tracking-tight mb-2">
                  Branch Feedback
                </h3>
                <p className="text-slate-500 font-medium">
                  Key survey insights and resolutions.
                </p>
              </div>

              <div className="space-y-4">
                {cxFeedbackLog.map((log, idx) => (
                  <div
                    key={idx}
                    className="bg-slate-50 border border-slate-200 rounded-2xl overflow-hidden shadow-sm hover:shadow-md transition-shadow"
                  >
                    <div className="bg-slate-100/50 px-6 py-3 border-b border-slate-200 font-black text-slate-700 flex items-center">
                      <MessageSquare className="w-5 h-5 mr-3 text-slate-400" />
                      GSF {log.site}
                    </div>
                    <div className="p-6 grid grid-cols-1 md:grid-cols-2 gap-8">
                      <div>
                        <p className="text-xs font-black text-rose-500 uppercase tracking-wider mb-2">
                          Branch Comment
                        </p>
                        <p className="text-slate-600 font-medium text-sm leading-relaxed italic">
                          "{log.issue}"
                        </p>
                      </div>
                      <div>
                        <p className="text-xs font-black text-blue-600 uppercase tracking-wider mb-2">
                          Resolution
                        </p>
                        <p className="text-slate-700 font-medium text-sm leading-relaxed">
                          {log.resolution}
                        </p>
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}
      </main>

      {/* Global CSS for Tailwind Animations */}
      <style
        dangerouslySetInnerHTML={{
          __html: `
        @keyframes fadeIn {
          from { opacity: 0; transform: translateY(15px); }
          to { opacity: 1; transform: translateY(0); }
        }
      `,
        }}
      />

      {/* MODALS */}
      {showAuthModal && (
        <div className="fixed inset-0 bg-slate-900/50 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-3xl shadow-2xl w-full max-w-md overflow-hidden animate-fade-in-up">
            <div className="p-6 bg-slate-50 border-b border-slate-100 flex justify-between items-center">
              <h3 className="text-xl font-black text-slate-800 flex items-center">
                <Lock className="w-5 h-5 mr-2 text-slate-500" /> Admin Access
              </h3>
              <button
                onClick={() => setShowAuthModal(false)}
                className="text-slate-400 hover:text-slate-600"
              >
                <X className="w-5 h-5" />
              </button>
            </div>
            <div className="p-6">
              <p className="text-sm font-medium text-slate-500 mb-4">
                Enter password to enable data editing.
              </p>
              <input
                type="password"
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                placeholder="Password (admin123)"
                className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:outline-none focus:ring-2 focus:ring-blue-500 mb-6"
                onKeyDown={(e) => e.key === "Enter" && handleLogin()}
              />
              <button
                onClick={handleLogin}
                className="w-full py-3 bg-blue-600 hover:bg-blue-700 text-white rounded-xl font-bold transition-colors shadow-md"
              >
                Unlock
              </button>
            </div>
          </div>
        </div>
      )}

      {showAddModal && (
        <div className="fixed inset-0 bg-slate-900/50 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-3xl shadow-2xl w-full max-w-3xl max-h-[90vh] overflow-hidden flex flex-col animate-fade-in-up">
            <div className="p-6 bg-slate-50 border-b border-slate-100 flex justify-between items-center">
              <h3 className="text-xl font-black text-slate-800 flex items-center">
                <Plus className="w-5 h-5 mr-2 text-blue-500" /> Data Editor
              </h3>
              <button
                onClick={() => {
                  setShowAddModal(false);
                  handleCancelEdit();
                }}
                className="text-slate-400 hover:text-slate-600"
              >
                <X className="w-5 h-5" />
              </button>
            </div>
            <div className="p-6 overflow-y-auto flex-1">
              <div className="mb-6">
                <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-2">
                  Select Dataset
                </label>
                <select
                  value={selectedTable}
                  onChange={(e) => {
                    setSelectedTable(e.target.value as any);
                    handleCancelEdit();
                  }}
                  className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:outline-none focus:ring-2 focus:ring-blue-500 bg-white"
                >
                  {Object.keys(data).map((key) => (
                    <option key={key} value={key}>
                      {key}
                    </option>
                  ))}
                </select>
              </div>

              <div className="my-4 space-y-2 max-h-[200px] overflow-y-auto pr-2">
                <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-2">
                  Existing Entries
                </label>
                {(data[selectedTable] || []).map((item, index) => (
                  <div
                    key={index}
                    className={`flex items-center justify-between p-3 rounded-lg border ${
                      editingIndex === index
                        ? "bg-blue-50 border-blue-200"
                        : "bg-slate-50 border-slate-200"
                    }`}
                  >
                    <p
                      className="text-xs text-slate-600 font-mono truncate w-full pr-2"
                      title={Object.values(item).join(" | ")}
                    >
                      {Object.values(item).join(" | ")}
                    </p>
                    <div className="flex space-x-2 flex-shrink-0">
                      <button
                        onClick={() => handleEditClick(index)}
                        className="font-bold text-xs text-blue-600 hover:underline"
                      >
                        EDIT
                      </button>
                      <button
                        onClick={() => handleDeleteClick(index)}
                        className="font-bold text-xs text-rose-600 hover:underline"
                      >
                        DEL
                      </button>
                    </div>
                  </div>
                ))}
              </div>

              <div className="space-y-4 border-t border-slate-200 pt-4 mt-4">
                <div className="flex justify-between items-center mb-2">
                  <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider">
                    {editingIndex !== null
                      ? `Editing Entry #${editingIndex + 1}`
                      : "Add New Entry"}
                  </label>
                  {editingIndex !== null && (
                    <button
                      onClick={handleCancelEdit}
                      className="text-xs font-bold text-slate-500 hover:underline"
                    >
                      Cancel Edit & Add New
                    </button>
                  )}
                </div>
                {data[selectedTable] && data[selectedTable].length > 0 ? (
                  Object.keys(data[selectedTable][0]).map((key) => (
                    <div key={key}>
                      <label className="block text-sm font-medium text-slate-700 mb-1 capitalize">
                        {key.replace(/([A-Z])/g, " $1")}
                      </label>
                      <input
                        type={
                          typeof data[selectedTable][0][key] === "number"
                            ? "number"
                            : "text"
                        }
                        value={formData[key] || ""}
                        onChange={(e) =>
                          setFormData({ ...formData, [key]: e.target.value })
                        }
                        className="w-full px-4 py-2 rounded-xl border border-slate-200 focus:outline-none focus:ring-2 focus:ring-blue-500"
                        placeholder={`Enter ${key}...`}
                        disabled={
                          key === "total" && selectedTable === "spendData"
                        }
                      />
                    </div>
                  ))
                ) : (
                  <p className="text-sm text-slate-500 italic">
                    This dataset is empty. Add the first entry.
                  </p>
                )}
              </div>
            </div>
            <div className="p-6 bg-slate-50 border-t border-slate-100 flex justify-end space-x-3">
              <button
                onClick={() => {
                  setShowAddModal(false);
                  handleCancelEdit();
                }}
                className="px-6 py-2.5 rounded-xl font-bold text-slate-600 hover:bg-slate-200 transition-colors"
              >
                Cancel
              </button>
              <button
                onClick={handleSaveData}
                className="px-6 py-2.5 bg-blue-600 hover:bg-blue-700 text-white rounded-xl font-bold transition-colors shadow-md"
              >
                {editingIndex !== null ? "Save Changes" : "Save New Data"}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
