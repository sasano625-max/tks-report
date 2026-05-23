"use client";

import { useState, useEffect } from "react";
import { saveAs } from "file-saver";
import { generatePPT, ReportData } from "@/lib/pptx-generator";
import { supabase } from "@/lib/supabase";
import { cn } from "@/lib/utils";
import { 
  Camera, 
  FileText, 
  LayoutDashboard, 
  Plus, 
  Send, 
  Download,
  Trash2,
  CheckCircle2,
  ChevronRight,
  Users,
  ShieldCheck,
  LogOut,
  Clock,
  UploadCloud
} from "lucide-react";
import { motion, AnimatePresence } from "framer-motion";

type Role = "vendor" | "admin";

export type PPTLayoutConfig = {
  tableX: number; tableY: number; tableW1: number; tableW2: number; tableRowH: number; tableFontSize: number;
  beforeX: number; beforeY: number; beforeW: number; beforeH: number;
  afterX: number; afterY: number; afterW: number; afterH: number;
  frontX: number; frontY: number; frontW: number; frontH: number;
  sideLeftX: number; sideLeftY: number; sideLeftW: number; sideLeftH: number;
  sideRightX: number; sideRightY: number; sideRightW: number; sideRightH: number;
  other1X: number; other1Y: number; other1W: number; other1H: number;
  other2X: number; other2Y: number; other2W: number; other2H: number;
};

export const DEFAULT_LAYOUT: PPTLayoutConfig = {
  tableX: 0.2, tableY: 0.35, tableW1: 1.0, tableW2: 3.2, tableRowH: 0.22, tableFontSize: 9,
  beforeX: 1.1, beforeY: 2.3, beforeW: 4.2, beforeH: 5.0,
  afterX: 8.0, afterY: 2.3, afterW: 4.2, afterH: 5.0,
  frontX: 0.2, frontY: 2.3, frontW: 4.2, frontH: 5.0,
  sideLeftX: 4.55, sideLeftY: 2.3, sideLeftW: 4.2, sideLeftH: 5.0,
  sideRightX: 8.9, sideRightY: 2.3, sideRightW: 4.2, sideRightH: 5.0,
  other1X: 1.1, other1Y: 2.3, other1W: 4.2, other1H: 5.0,
  other2X: 8.0, other2Y: 2.3, other2W: 4.2, other2H: 5.0,
};

export default function Home() {
  const [role, setRole] = useState<Role | null>(null);
  const [submissions, setSubmissions] = useState<ReportData[]>([]);
  const [loadingSubmissions, setLoadingSubmissions] = useState(false);
  const [templateUrl, setTemplateUrl] = useState<string | null>(null);
  const [isUploadingTemplate, setIsUploadingTemplate] = useState(false);
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());
  const [isBatchDownloading, setIsBatchDownloading] = useState(false);
  const [isUploadingCSV, setIsUploadingCSV] = useState(false);
  const [plannedStores, setPlannedStores] = useState<any[]>([]);
  const [searchQuery, setSearchQuery] = useState("");
  const [filterStatus, setFilterStatus] = useState<"all" | "submitted" | "waiting">("all");
  const [editingSubId, setEditingSubId] = useState<string | null>(null);
  const [editData, setEditData] = useState<any>(null);
  
  // Vendor Form State
  const [formData, setFormData] = useState<ReportData>({
    storeName: "",
    completionDate: new Date().toISOString().split("T")[0],
    type: "Samsung Brand Table",
    monitorLeft: "",
    monitorRight: "",
    images: {}
  });

  const [isSubmitting, setIsSubmitting] = useState(false);
  const [showSuccess, setShowSuccess] = useState(false);

  const [isAuthenticated, setIsAuthenticated] = useState(false);
  const [adminPassword, setAdminPassword] = useState("");
  const [showAdminAuth, setShowAdminAuth] = useState(false);
  const [authError, setAuthError] = useState(false);

  // Delete All State
  const [showDeleteAllAuth, setShowDeleteAllAuth] = useState(false);
  const [deleteAllPassword, setDeleteAllPassword] = useState("");
  const [deleteAllError, setDeleteAllError] = useState(false);
  const [isDeletingAll, setIsDeletingAll] = useState(false);

  // PPT Layout Settings State
  const [showLayoutSettings, setShowLayoutSettings] = useState(false);
  const [pptLayout, setPptLayout] = useState<PPTLayoutConfig>(DEFAULT_LAYOUT);

  useEffect(() => {
    if (typeof window !== "undefined") {
      const saved = localStorage.getItem("pptLayout");
      if (saved) {
        try {
          setPptLayout(JSON.parse(saved));
        } catch (e) {
          console.error("Failed to parse pptLayout");
        }
      }
    }
  }, []);

  const handleSavePptLayout = (newLayout: PPTLayoutConfig) => {
    setPptLayout(newLayout);
    localStorage.setItem("pptLayout", JSON.stringify(newLayout));
    setShowLayoutSettings(false);
  };

  // Fetch submissions on admin role switch
  useEffect(() => {
    if (role === "admin" && !isAuthenticated) {
      setShowAdminAuth(true);
      setRole(null); // 一旦ロールをリセットして認証を待つ
    }
    if (role === "admin" && isAuthenticated) {
      fetchSubmissions();
    }
  }, [role, isAuthenticated]);

  const handleAdminAuth = (e: React.FormEvent) => {
    e.preventDefault();
    if (adminPassword === "tkstks") {
      setIsAuthenticated(true);
      setRole("admin");
      setShowAdminAuth(false);
      setAuthError(false);
    } else {
      setAuthError(true);
      setAdminPassword("");
    }
  };

  const fetchSubmissions = async () => {
    setLoadingSubmissions(true);
    const { data, error } = await supabase
      .from("reports")
      .select("*")
      .order("created_at", { ascending: false });

    if (error) {
      console.error("Error fetching submissions:", error);
    } else {
      setSubmissions(data || []);
    }
    setLoadingSubmissions(false);
  };

  const fetchPlannedStores = async () => {
    const { data, error } = await supabase
      .from("reports")
      .select("*")
      .is("images", null)
      .order("completion_date", { ascending: true });
    
    if (!error && data) {
      setPlannedStores(data);
    }
  };

  useEffect(() => {
    if (role === "admin" || role === "vendor") {
      fetchPlannedStores();
    }
  }, [role]);

  const fetchTemplate = async () => {
    const { data: { publicUrl } } = supabase.storage
      .from("reports")
      .getPublicUrl("assets/template.jpg");
    
    // Check if exists by trying to fetch it
    try {
      const resp = await fetch(publicUrl, { method: 'HEAD' });
      if (resp.ok) setTemplateUrl(publicUrl);
    } catch (e) {
      console.log("No template found");
    }
  };

  useEffect(() => {
    fetchTemplate();
  }, []);

  const handleImageUpload = (e: React.ChangeEvent<HTMLInputElement>, key: keyof ReportData["images"]) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      const img = new Image();
      img.onload = () => {
        const MAX_WIDTH = 1280;
        const MAX_HEIGHT = 1280;
        let width = img.width;
        let height = img.height;

        if (width > height) {
          if (width > MAX_WIDTH) {
            height = Math.round((height * MAX_WIDTH) / width);
            width = MAX_WIDTH;
          }
        } else {
          if (height > MAX_HEIGHT) {
            width = Math.round((width * MAX_HEIGHT) / height);
            height = MAX_HEIGHT;
          }
        }

        const canvas = document.createElement("canvas");
        canvas.width = width;
        canvas.height = height;
        const ctx = canvas.getContext("2d");
        
        if (ctx) {
          ctx.drawImage(img, 0, 0, width, height);
          // Compress to JPEG with 70% quality to reduce Supabase load
          const dataUrl = canvas.toDataURL("image/jpeg", 0.7);
          
          setFormData(prev => ({
            ...prev,
            images: { ...prev.images, [key]: dataUrl }
          }));
        }
      };
      img.src = event.target?.result as string;
    };
    reader.readAsDataURL(file);
  };

  const uploadImage = async (base64: string, path: string) => {
    const blob = await (await fetch(base64)).blob();
    // Use a simpler path to avoid encoding issues
    const { data, error } = await supabase.storage
      .from("reports")
      .upload(path, blob, { upsert: true, contentType: 'image/jpeg' });
    
    if (error) throw error;
    
    const { data: { publicUrl } } = supabase.storage
      .from("reports")
      .getPublicUrl(path);
    
    return publicUrl;
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!formData.storeName) return;
    
    setIsSubmitting(true);
    
    try {
      const timestamp = Date.now();
      const uploadedImages: Record<string, string> = {};

      // Upload images to Supabase Storage
      for (const [key, base64] of Object.entries(formData.images)) {
        if (base64) {
          // Use only ASCII for the path
          const safeStoreName = "store_" + timestamp; 
          const path = `${safeStoreName}/${key}.jpg`;
          const url = await uploadImage(base64, path);
          uploadedImages[key] = url;
        }
      }

      // Save report data to Supabase Database
      const { error } = await supabase
        .from("reports")
        .insert([{
          store_name: formData.storeName,
          completion_date: formData.completionDate,
          type: formData.type,
          monitor_left: formData.monitorLeft,
          monitor_right: formData.monitorRight,
          images: uploadedImages,
          created_at: new Date().toISOString()
        }]);

      if (error) throw error;

      setShowSuccess(true);
      setTimeout(() => setShowSuccess(false), 3000);
      
      // Force refresh submissions if in admin mode or just to be safe
      fetchSubmissions();

      // Reset form
      setFormData({
        storeName: "",
        completionDate: new Date().toISOString().split("T")[0],
        type: "Samsung Brand Table",
        monitorLeft: "",
        monitorRight: "",
        images: {}
      });

    } catch (error) {
      console.error("Submission failed:", error);
      alert("送信に失敗しました。詳細: " + (error as Error).message);
    } finally {
      setIsSubmitting(false);
    }
  };

  const [downloadingId, setDownloadingId] = useState<string | null>(null);

  const handleDownload = async (rawSub: any, id: string) => {
    setDownloadingId(id);
    try {
      // Map Supabase snake_case to ReportData camelCase
      const data: ReportData = {
        storeName: rawSub.store_name || rawSub.storeName || "名称未設定",
        completionDate: rawSub.completion_date || rawSub.completionDate || "",
        type: rawSub.type || "",
        monitorLeft: rawSub.monitor_left || rawSub.monitorLeft || "",
        monitorRight: rawSub.monitor_right || rawSub.monitorRight || "",
        images: rawSub.images || {}
      };

      // Server-side generation call
      const response = await fetch("/api/generate-ppt", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ ...data, templateUrl, layoutConfig: pptLayout }),
      });

      if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        throw new Error(errorData.error || "サーバーでのパワポ生成に失敗しました");
      }

      const blob = await response.blob();
      console.log("Blob received from server, size:", blob.size);

      if (blob.size < 100) {
        throw new Error("生成されたファイルが空か、無効な形式です。");
      }

      // Use file-saver's saveAs for maximum reliability
      const filename = `${data.storeName}_${new Date().toISOString().split('T')[0]}.pptx`;
      saveAs(blob, filename);
      
      console.log("PPT download triggered via file-saver:", filename);
      console.log("PPT generation finished successfully via Server API");
    } catch (error) {
      console.error("PPT generation failed", error);
      alert("パワポの作成に失敗しました。\nエラー詳細: " + (error as Error).message);
    } finally {
      setDownloadingId(null);
    }
  };

  const [deletingId, setDeletingId] = useState<string | null>(null);

  const handleDelete = async (id: string) => {
    if (!confirm("この報告を削除してもよろしいですか？この操作は取り消せません。")) return;
    
    setDeletingId(id);
    try {
      const { error } = await supabase
        .from("reports")
        .delete()
        .eq("id", id);

      if (error) throw error;
      
      setSubmissions(prev => prev.filter(sub => (sub as any).id !== id));
      setSelectedIds(prev => {
        const next = new Set(prev);
        next.delete(id);
        return next;
      });
      
      console.log("Report deleted successfully:", id);
    } catch (error) {
      console.error("Delete failed:", error);
      alert("削除に失敗しました。\n詳細: " + (error as Error).message);
    } finally {
      setDeletingId(null);
    }
  };

  const handleDeleteAll = async (e: React.FormEvent) => {
    e.preventDefault();
    if (deleteAllPassword !== "tkstks") {
      setDeleteAllError(true);
      setDeleteAllPassword("");
      return;
    }

    if (!confirm("本当に全てのデータを削除しますか？\nこの操作は取り消せません。")) {
      setShowDeleteAllAuth(false);
      setDeleteAllPassword("");
      setDeleteAllError(false);
      return;
    }

    setIsDeletingAll(true);
    try {
      // Deleting all rows by using a condition that matches all (neq invalid UUID)
      const { error } = await supabase
        .from("reports")
        .delete()
        .neq("id", "00000000-0000-0000-0000-000000000000");

      if (error) throw error;

      setSubmissions([]);
      setSelectedIds(new Set());
      setPlannedStores([]);
      alert("全データを削除しました");
      
      setShowDeleteAllAuth(false);
      setDeleteAllPassword("");
      setDeleteAllError(false);
    } catch (error) {
      alert("削除に失敗しました: " + (error as Error).message);
    } finally {
      setIsDeletingAll(false);
    }
  };

  const handleBatchDelete = async () => {
    if (selectedIds.size === 0) return;
    if (!confirm(`${selectedIds.size}件のデータを削除してもよろしいですか？`)) return;

    setIsBatchDownloading(true); // Reuse loading state or add new one
    try {
      const { error } = await supabase
        .from("reports")
        .delete()
        .in("id", Array.from(selectedIds));

      if (error) throw error;

      setSubmissions(prev => prev.filter(sub => !selectedIds.has((sub as any).id)));
      setSelectedIds(new Set());
      alert("選択したデータを削除しました");
    } catch (error) {
      alert("削除に失敗しました: " + (error as Error).message);
    } finally {
      setIsBatchDownloading(false);
    }
  };

  const filteredSubmissions = submissions.filter(sub => {
    const storeName = (sub as any).store_name || sub.storeName || "";
    const matchesSearch = storeName.toLowerCase().includes(searchQuery.toLowerCase());
    
    if (filterStatus === "submitted") return matchesSearch && !!sub.images;
    if (filterStatus === "waiting") return matchesSearch && !sub.images;
    return matchesSearch;
  });

  const toggleSelectAll = () => {
    if (selectedIds.size === filteredSubmissions.length) {
      setSelectedIds(new Set());
    } else {
      setSelectedIds(new Set(filteredSubmissions.map(s => (s as any).id)));
    }
  };

  const handleEditStart = (sub: any) => {
    setEditingSubId(sub.id);
    setEditData({
      store_name: sub.store_name || sub.storeName,
      completion_date: sub.completion_date || sub.completionDate,
      type: sub.type,
      monitor_left: sub.monitor_left || sub.monitorLeft,
      monitor_right: sub.monitor_right || sub.monitorRight
    });
  };

  const handleEditSave = async () => {
    if (!editingSubId) return;
    try {
      const { error } = await supabase
        .from("reports")
        .update(editData)
        .eq("id", editingSubId);

      if (error) throw error;
      
      setSubmissions(prev => prev.map(s => (s as any).id === editingSubId ? { ...s, ...editData } : s));
      setEditingSubId(null);
    } catch (error) {
      alert("更新に失敗しました: " + (error as Error).message);
    }
  };

  const handleBatchDownload = async () => {
    if (selectedIds.size === 0) return;
    
    setIsBatchDownloading(true);
    try {
      const selectedReports = submissions.filter(sub => selectedIds.has((sub as any).id));
      
      const response = await fetch("/api/generate-ppt", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ reports: selectedReports, templateUrl, layoutConfig: pptLayout }),
      });

      if (!response.ok) throw new Error("一括ダウンロードに失敗しました");

      const blob = await response.blob();
      saveAs(blob, `Batch_Report_${new Date().toISOString().split('T')[0]}.pptx`);
      
      setSelectedIds(new Set());
    } catch (error) {
      alert("エラー: " + (error as Error).message);
    } finally {
      setIsBatchDownloading(false);
    }
  };

  const handleCSVUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsUploadingCSV(true);
    const reader = new FileReader();
    
    reader.onload = async (event) => {
      try {
        const buffer = event.target?.result as ArrayBuffer;
        let text = "";
        
        try {
          // Attempt UTF-8 with fatal:true to catch encoding errors
          const utf8Decoder = new TextDecoder("utf-8", { fatal: true });
          text = utf8Decoder.decode(buffer);
        } catch (e) {
          // Fallback to Shift-JIS if UTF-8 fails
          const sjisDecoder = new TextDecoder("shift-jis");
          text = sjisDecoder.decode(buffer);
        }

        // If it's still looking like it might be wrong (e.g. no delimiters found), 
        // it might be UTF-8 but without BOM or something, but usually the fatal:true covers it.
        if (!text.includes(",") && !text.includes("\t")) {
           // One more try with Shift-JIS just in case UTF-8 succeeded but produced garbage
           const sjisDecoder = new TextDecoder("shift-jis");
           text = sjisDecoder.decode(buffer);
        }

        const lines = text.split(/\r?\n/).filter(line => line.trim());
        if (lines.length === 0) throw new Error("ファイルが空です");

        // Basic delimiter detection
        const delimiter = lines[0].includes("\t") ? "\t" : ",";
        
        // Detect header (if it contains '店舗名' or '予定日')
        const hasHeader = lines[0].includes("店舗名") || lines[0].includes("予定日");
        const dataLines = hasHeader ? lines.slice(1) : lines;

        const newRows = dataLines.map((line, idx) => {
          const parts = line.split(delimiter).map(p => p.trim().replace(/^"|"$/g, '')); 
          if (parts.length < 4) return null; // Need at least up to column 3 (Planned Date)
          
          // Planned Date is at index 3
          let dateStr = parts[3];
          if (!dateStr) return null;
          
          // Date Normalization (Supports YYYY-MM-DD, YYYY/MM/DD, YY-MM-DD, etc.)
          dateStr = dateStr.replace(/\//g, "-");
          const dateParts = dateStr.split("-");
          if (dateParts.length !== 3) return null;
          
          let y = dateParts[0];
          if (y.length === 2) y = "20" + y; // Assume 20xx for 2-digit year
          const m = dateParts[1].padStart(2, '0');
          const d = dateParts[2].padStart(2, '0');
          const date = `${y}-${m}-${d}`;

          return {
            store_name: parts[0],
            completion_date: date,
            type: "Brand Table", // default type as it's not in the new CSV
            images: null,
            created_at: new Date().toISOString()
          };
        }).filter(Boolean);

        if (newRows.length === 0) {
          throw new Error(`有効なデータが見つかりませんでした。\n形式: 店舗名,住所,電話番号,予定日...\n読み取った最初の行: ${lines[0]}`);
        }

        const { error } = await supabase.from("reports").insert(newRows as any);
        if (error) {
          console.error("Supabase Insert Error:", error);
          throw new Error(`データベース登録エラー: ${error.message} (${error.details || ''})`);
        }

        alert(`${newRows.length}件の店舗を登録しました`);
        fetchSubmissions();
        fetchPlannedStores();
      } catch (error) {
        console.error("CSV Parse Error:", error);
        alert("CSV読み込みエラー: " + (error as Error).message);
      } finally {
        setIsUploadingCSV(false);
        e.target.value = "";
      }
    };

    reader.readAsArrayBuffer(file);
  };

  return (
    <div className="min-h-screen bg-slate-50 pb-24">
      {/* Admin Auth Modal */}
      <AnimatePresence>
        {showAdminAuth && (
          <motion.div 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 z-[100] flex items-center justify-center p-6 bg-slate-900/60 backdrop-blur-sm"
          >
            <motion.div 
              initial={{ scale: 0.9, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.9, opacity: 0 }}
              className="w-full max-w-sm bg-white rounded-[32px] p-10 shadow-2xl border border-slate-100"
            >
              <div className="flex flex-col items-center">
                <div className="w-16 h-16 bg-indigo-50 rounded-2xl flex items-center justify-center mb-6">
                  <ShieldCheck className="text-indigo-600" size={32} />
                </div>
                <h2 className="text-2xl font-black text-slate-900 mb-2 text-center">管理者認証</h2>
                <p className="text-sm font-bold text-slate-400 mb-8 text-center uppercase tracking-widest">管理者モードのアクセス確認</p>
                
                <form onSubmit={handleAdminAuth} className="w-full space-y-6">
                  <div className="space-y-2">
                    <input 
                      type="password" 
                      value={adminPassword}
                      onChange={(e) => {
                        setAdminPassword(e.target.value);
                        setAuthError(false);
                      }}
                      placeholder="パスワードを入力"
                      className={`w-full bg-slate-50 border ${authError ? 'border-red-500 ring-4 ring-red-500/10' : 'border-slate-200 focus:border-indigo-500 focus:ring-4 focus:ring-indigo-500/5'} rounded-2xl p-4 text-center font-bold text-slate-700 outline-none transition-all`}
                      autoFocus
                    />
                    {authError && (
                      <p className="text-[10px] font-black text-red-500 text-center uppercase tracking-widest">パスワードが正しくありません</p>
                    )}
                  </div>
                  
                  <div className="flex gap-4">
                    <button 
                      type="button"
                      onClick={() => {
                        setShowAdminAuth(false);
                        setAdminPassword("");
                        setAuthError(false);
                      }}
                      className="flex-1 py-4 bg-slate-100 hover:bg-slate-200 text-slate-500 rounded-2xl font-black transition-all active:scale-[0.98]"
                    >
                      戻る
                    </button>
                    <button 
                      type="submit"
                      className="flex-[2] py-4 bg-indigo-600 hover:bg-indigo-700 text-white rounded-2xl font-black shadow-xl shadow-indigo-500/20 transition-all active:scale-[0.98]"
                    >
                      認証
                    </button>
                  </div>
                </form>
              </div>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Delete All Data Modal */}
      <AnimatePresence>
        {showDeleteAllAuth && (
          <motion.div 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 z-[100] flex items-center justify-center p-6 bg-slate-900/60 backdrop-blur-sm"
          >
            <motion.div 
              initial={{ scale: 0.9, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.9, opacity: 0 }}
              className="w-full max-w-sm bg-white rounded-[32px] p-10 shadow-2xl border border-slate-100"
            >
              <div className="flex flex-col items-center">
                <div className="w-16 h-16 bg-red-50 rounded-2xl flex items-center justify-center mb-6">
                  <Trash2 className="text-red-600" size={32} />
                </div>
                <h2 className="text-2xl font-black text-slate-900 mb-2 text-center">全データ初期化</h2>
                <p className="text-sm font-bold text-slate-400 mb-8 text-center uppercase tracking-widest">削除パスワードを入力</p>
                
                <form onSubmit={handleDeleteAll} className="w-full space-y-6">
                  <div className="space-y-2">
                    <input 
                      type="password" 
                      value={deleteAllPassword}
                      onChange={(e) => {
                        setDeleteAllPassword(e.target.value);
                        setDeleteAllError(false);
                      }}
                      placeholder="パスワードを入力"
                      className={`w-full bg-slate-50 border ${deleteAllError ? 'border-red-500 ring-4 ring-red-500/10' : 'border-slate-200 focus:border-red-500 focus:ring-4 focus:ring-red-500/5'} rounded-2xl p-4 text-center font-bold text-slate-700 outline-none transition-all`}
                      autoFocus
                    />
                    {deleteAllError && (
                      <p className="text-[10px] font-black text-red-500 text-center uppercase tracking-widest">パスワードが正しくありません</p>
                    )}
                  </div>
                  
                  <div className="flex gap-4">
                    <button 
                      type="button"
                      onClick={() => {
                        setShowDeleteAllAuth(false);
                        setDeleteAllPassword("");
                        setDeleteAllError(false);
                      }}
                      className="flex-1 py-4 bg-slate-100 hover:bg-slate-200 text-slate-500 rounded-2xl font-black transition-all active:scale-[0.98]"
                    >
                      戻る
                    </button>
                    <button 
                      type="submit"
                      disabled={isDeletingAll}
                      className="flex-[2] py-4 bg-red-600 hover:bg-red-700 text-white rounded-2xl font-black shadow-xl shadow-red-500/20 transition-all active:scale-[0.98] disabled:opacity-50"
                    >
                      {isDeletingAll ? "削除中..." : "実行"}
                    </button>
                  </div>
                </form>
              </div>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Header */}
      {role && (
        <header className="glass sticky top-0 z-50 px-6 py-4 flex items-center justify-between shadow-sm">
          <nav className="flex justify-between items-center w-full">
            <div className="flex items-center gap-3">
              <img src="/logo.png" alt="Logo" className="w-8 h-8 object-contain" />
              <div className="flex flex-col">
                <h1 className="text-xl font-black bg-clip-text text-transparent bg-gradient-to-r from-blue-900 to-blue-600 tracking-tight">
                  RepoMaker
                </h1>
                <span className="text-[8px] text-slate-400 font-bold tracking-[0.2em] uppercase">
                  Automated Report System
                </span>
              </div>
            </div>
            
            <button 
              onClick={() => setRole(null)}
              className="flex items-center gap-2 px-3 py-1.5 rounded-lg text-slate-400 hover:text-red-500 hover:bg-red-50 transition-all text-xs font-bold"
            >
              <LogOut className="w-4 h-4" />
              <span>終了</span>
            </button>
          </nav>
        </header>
      )}

      <main className={cn("max-w-md mx-auto p-6", !role && "flex flex-col justify-center min-h-screen")}>
        <AnimatePresence mode="wait">
          {!role ? (
            <motion.div
              key="selection"
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 1.05 }}
              className="space-y-12"
            >
              <div className="text-center space-y-6">
                <div className="relative inline-block">
                  <img src="/logo.png" alt="TKS RepoMaker" className="w-48 mx-auto" />
                </div>
                <div className="space-y-2">
                  <p className="text-slate-500 font-medium">利用する権限を選択してください</p>
                </div>
              </div>

              <div className="grid gap-4">
                <button 
                  onClick={() => setRole("vendor")}
                  className="group relative overflow-hidden bg-white p-8 rounded-3xl shadow-xl shadow-slate-200/50 border border-slate-100 flex flex-col items-center gap-4 transition-all hover:border-blue-500 hover:shadow-blue-100 hover:-translate-y-1 active:scale-95"
                >
                  <div className="w-16 h-16 bg-blue-50 text-blue-600 rounded-2xl flex items-center justify-center group-hover:bg-blue-600 group-hover:text-white transition-all">
                    <Users className="w-8 h-8" />
                  </div>
                  <div className="text-center">
                    <h3 className="text-xl font-bold text-slate-900">業者用</h3>
                    <p className="text-sm text-slate-400 mt-1">設置報告・写真アップロード</p>
                  </div>
                  <ChevronRight className="absolute right-6 top-1/2 -translate-y-1/2 w-6 h-6 text-slate-200 group-hover:text-blue-500 group-hover:translate-x-1 transition-all" />
                </button>

                <button 
                  onClick={() => setRole("admin")}
                  className="group relative overflow-hidden bg-white p-8 rounded-3xl shadow-xl shadow-slate-200/50 border border-slate-100 flex flex-col items-center gap-4 transition-all hover:border-indigo-500 hover:shadow-indigo-100 hover:-translate-y-1 active:scale-95"
                >
                  <div className="w-16 h-16 bg-indigo-50 text-indigo-600 rounded-2xl flex items-center justify-center group-hover:bg-indigo-600 group-hover:text-white transition-all">
                    <ShieldCheck className="w-8 h-8" />
                  </div>
                  <div className="text-center">
                    <h3 className="text-xl font-bold text-slate-900">管理者</h3>
                    <p className="text-sm text-slate-400 mt-1">報告確認・パワポ生成</p>
                  </div>
                  <ChevronRight className="absolute right-6 top-1/2 -translate-y-1/2 w-6 h-6 text-slate-200 group-hover:text-indigo-500 group-hover:translate-x-1 transition-all" />
                </button>
              </div>

              <p className="text-center text-[10px] text-slate-300 font-bold tracking-widest uppercase">
                Field Reporting System v3.1
              </p>
            </motion.div>
          ) : role === "vendor" ? (
            <motion.div
              key="vendor-view"
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
              className="space-y-8"
            >
              <section className="space-y-4">
                <div className="flex items-center gap-2 text-slate-800 font-semibold">
                  <FileText className="w-5 h-5 text-blue-500" />
                  <h2>店舗・設置情報</h2>
                </div>
                
                <div className="grid gap-4 bg-white p-6 rounded-2xl shadow-sm border border-slate-100">
                  <div className="space-y-2">
                    <label className="text-xs font-bold text-slate-400 uppercase tracking-wider">設置日</label>
                    <input 
                      type="date" 
                      className="w-full px-4 py-3 rounded-xl bg-slate-50 border-transparent outline-none focus:bg-white focus:ring-2 focus:ring-blue-500/20 transition-all"
                      value={formData.completionDate}
                      onChange={e => setFormData({...formData, completionDate: e.target.value})}
                    />
                  </div>

                  {/* Planned Store Selector */}
                  <div className="space-y-2 bg-blue-50/50 p-4 rounded-2xl border border-blue-100/50">
                    <label className="text-[10px] font-bold text-blue-400 uppercase tracking-wider">予定店舗から選択</label>
                    <select 
                      className="w-full px-4 py-3 rounded-xl bg-white border-slate-200 outline-none text-sm font-medium"
                      onChange={(e) => {
                        const store = plannedStores.find(s => s.id === e.target.value);
                        if (store) {
                          setFormData({
                            ...formData,
                            storeName: store.store_name,
                            type: store.type || "Brand Table"
                          });
                        }
                      }}
                    >
                      <option value="">-- 対象店舗を選択してください --</option>
                      {plannedStores
                        .filter(s => s.completion_date === formData.completionDate)
                        .map(s => (
                        <option key={s.id} value={s.id}>{s.store_name}</option>
                      ))}
                    </select>
                    {plannedStores.filter(s => s.completion_date === formData.completionDate).length === 0 && (
                      <p className="text-[10px] text-red-400 font-medium mt-1">※この日付の予定店舗はありません</p>
                    )}
                    <p className="text-[10px] text-blue-400 font-medium italic mt-1">※リストにない場合は以下に直接入力してください</p>
                  </div>

                  <div className="space-y-2">
                    <label className="text-xs font-bold text-slate-400 uppercase tracking-wider">店舗名</label>
                    <input 
                      type="text" 
                      placeholder="例: AMむさし村山店"
                      className="w-full px-4 py-3 rounded-xl bg-slate-50 border-transparent focus:bg-white focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 outline-none transition-all"
                      value={formData.storeName}
                      onChange={e => setFormData({...formData, storeName: e.target.value})}
                    />
                  </div>
                  
                  <div className="space-y-2">
                    <label className="text-xs font-bold text-slate-400 uppercase tracking-wider">タイプ</label>
                    <input 
                      type="text"
                      list="brand-types"
                      placeholder="例: Samsung"
                      className="w-full px-4 py-3 rounded-xl bg-slate-50 border-transparent outline-none focus:bg-white focus:ring-2 focus:ring-blue-500/20 transition-all"
                      value={formData.type}
                      onChange={e => setFormData({...formData, type: e.target.value})}
                    />
                    <datalist id="brand-types">
                      <option value="Samsung Brand Table" />
                      <option value="Google Brand Table" />
                      <option value="Apple Brand Table" />
                    </datalist>
                  </div>

                  <div className="grid grid-cols-2 gap-4">
                    <div className="space-y-2">
                      <label className="text-xs font-bold text-slate-400 uppercase tracking-wider">モニター番号 (左)</label>
                      <input 
                        type="number" 
                        placeholder="7"
                        className="w-full px-4 py-3 rounded-xl bg-slate-50 border-transparent outline-none focus:bg-white focus:ring-2 focus:ring-blue-500/20 transition-all"
                        value={formData.monitorLeft}
                        onChange={e => setFormData({...formData, monitorLeft: e.target.value})}
                      />
                    </div>
                    <div className="space-y-2">
                      <label className="text-xs font-bold text-slate-400 uppercase tracking-wider">モニター番号 (右)</label>
                      <input 
                        type="number" 
                        placeholder="8"
                        className="w-full px-4 py-3 rounded-xl bg-slate-50 border-transparent outline-none focus:bg-white focus:ring-2 focus:ring-blue-500/20 transition-all"
                        value={formData.monitorRight}
                        onChange={e => setFormData({...formData, monitorRight: e.target.value})}
                      />
                    </div>
                  </div>
                </div>
              </section>

              <section className="space-y-4">
                <div className="flex items-center gap-2 text-slate-800 font-semibold">
                  <Camera className="w-5 h-5 text-blue-500" />
                  <h2>報告写真アップロード</h2>
                </div>

                <div className="grid grid-cols-2 gap-4">
                  {[
                    { id: "before", label: "設置前" },
                    { id: "after", label: "設置後" },
                    { id: "front", label: "正面" },
                    { id: "sideLeft", label: "左側面" },
                    { id: "sideRight", label: "右側面" },
                    { id: "other1", label: "他社1" },
                    { id: "other2", label: "他社2" },
                  ].map((slot) => (
                    <div key={slot.id} className="relative aspect-square">
                      <label className="w-full h-full flex flex-col items-center justify-center border-2 border-dashed border-slate-200 rounded-2xl bg-white hover:bg-slate-50 transition-colors cursor-pointer overflow-hidden">
                        {formData.images[slot.id as keyof ReportData["images"]] ? (
                          <img 
                            src={formData.images[slot.id as keyof ReportData["images"]]} 
                            className="w-full h-full object-cover" 
                            alt={slot.label} 
                          />
                        ) : (
                          <>
                            <Plus className="w-8 h-8 text-slate-300 mb-2" />
                            <span className="text-xs font-bold text-slate-400">{slot.label}</span>
                          </>
                        )}
                        <input 
                          type="file" 
                          accept="image/*" 
                          className="hidden" 
                          onChange={(e) => handleImageUpload(e, slot.id as keyof ReportData["images"])}
                        />
                      </label>
                      {formData.images[slot.id as keyof ReportData["images"]] && (
                        <button 
                          onClick={() => setFormData(prev => ({
                            ...prev, 
                            images: { ...prev.images, [slot.id]: undefined }
                          }))}
                          className="absolute -top-2 -right-2 p-1.5 bg-red-500 text-white rounded-full shadow-lg"
                        >
                          <Trash2 className="w-3 h-3" />
                        </button>
                      )}
                    </div>
                  ))}
                </div>
              </section>

              <button 
                onClick={handleSubmit}
                disabled={isSubmitting || !formData.storeName}
                className={cn(
                  "w-full py-4 rounded-2xl font-bold flex items-center justify-center gap-2 shadow-xl transition-all",
                  isSubmitting || !formData.storeName 
                    ? "bg-slate-300 text-slate-500 cursor-not-allowed" 
                    : "bg-blue-600 text-white hover:bg-blue-700 active:scale-95"
                )}
              >
                {isSubmitting ? (
                  <div className="w-6 h-6 border-2 border-white/30 border-t-white rounded-full animate-spin" />
                ) : (
                  <>
                    <Send className="w-5 h-5" />
                    報告を送信する
                  </>
                )}
              </button>
            </motion.div>
          ) : (
            <motion.div
              key="admin-view"
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
              className="space-y-6"
            >
              <div className="flex items-center justify-between">
                <div className="flex items-center gap-2 text-slate-800 font-semibold">
                  <LayoutDashboard className="w-5 h-5 text-blue-500" />
                  <h2>報告一覧</h2>
                </div>
                <div className="flex items-center gap-2">
                  {selectedIds.size > 0 && (
                    <div className="flex items-center gap-2">
                      <button 
                        onClick={handleBatchDelete}
                        className="flex items-center gap-2 bg-red-50 text-red-600 px-3 py-2 rounded-xl text-xs font-bold border border-red-100 hover:bg-red-100 transition-all"
                      >
                        <Trash2 className="w-4 h-4" />
                        削除 ({selectedIds.size})
                      </button>
                      <button 
                        onClick={handleBatchDownload}
                        disabled={isBatchDownloading}
                        className="flex items-center gap-2 bg-blue-600 text-white px-4 py-2 rounded-xl text-xs font-bold shadow-lg hover:bg-blue-700 transition-all disabled:opacity-50"
                      >
                        {isBatchDownloading ? (
                          <div className="w-4 h-4 border-2 border-white/30 border-t-white rounded-full animate-spin" />
                        ) : (
                          <Download className="w-4 h-4" />
                        )}
                        出力 ({selectedIds.size})
                      </button>
                    </div>
                  )}
                </div>
              </div>

              {/* Admin Controls */}
              <div className="space-y-3">
                <div className="flex gap-2">
                  <div className="relative flex-1">
                    <input 
                      type="text" 
                      placeholder="店舗名で検索..."
                      className="w-full pl-10 pr-4 py-2.5 rounded-xl bg-white border border-slate-200 text-sm focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 outline-none transition-all"
                      value={searchQuery}
                      onChange={e => setSearchQuery(e.target.value)}
                    />
                    <Plus className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2 rotate-45" />
                  </div>
                </div>
                
                <div className="flex items-center justify-between">
                  <div className="flex p-1 bg-slate-100 rounded-xl">
                    {[
                      { id: "all", label: "すべて" },
                      { id: "submitted", label: "提出済" },
                      { id: "waiting", label: "未提出" }
                    ].map(btn => (
                      <button
                        key={btn.id}
                        onClick={() => setFilterStatus(btn.id as any)}
                        className={cn(
                          "px-4 py-1.5 rounded-lg text-[10px] font-bold transition-all",
                          filterStatus === btn.id 
                            ? "bg-white text-blue-600 shadow-sm" 
                            : "text-slate-500 hover:text-slate-700"
                        )}
                      >
                        {btn.label}
                      </button>
                    ))}
                  </div>
                  
                  <button 
                    onClick={toggleSelectAll}
                    className="text-[10px] font-bold text-blue-500 hover:underline px-2"
                  >
                    {selectedIds.size === filteredSubmissions.length && filteredSubmissions.length > 0 ? "全選択解除" : "全選択"}
                  </button>
                </div>
              </div>

              {/* Summary Stats Card */}
              <div className="grid grid-cols-2 gap-4">
                <div className="bg-white p-4 rounded-2xl border border-slate-100 shadow-sm flex items-center gap-3">
                  <div className="w-10 h-10 bg-blue-50 text-blue-600 rounded-xl flex items-center justify-center">
                    <CheckCircle2 className="w-5 h-5" />
                  </div>
                  <div>
                    <p className="text-[10px] text-slate-400 font-bold uppercase tracking-wider">全データ</p>
                    <p className="text-xl font-black text-slate-900">{submissions.length} <span className="text-xs font-normal text-slate-400">件</span></p>
                  </div>
                </div>
                
                <div className="bg-white p-4 rounded-2xl border border-slate-100 shadow-sm flex items-center gap-3 relative group overflow-hidden">
                  <div className="w-10 h-10 bg-indigo-50 text-indigo-600 rounded-xl flex items-center justify-center">
                    <UploadCloud className="w-5 h-5" />
                  </div>
                  <div>
                    <p className="text-[10px] text-slate-400 font-bold uppercase tracking-wider">CSV登録</p>
                    <p className="text-[10px] font-bold text-indigo-500 cursor-pointer hover:underline" title="形式: 日付, 店舗名, タイプ">店舗一括登録</p>
                  </div>
                  <input 
                    type="file" 
                    accept=".csv" 
                    className="absolute inset-0 opacity-0 cursor-pointer" 
                    onChange={handleCSVUpload}
                    disabled={isUploadingCSV}
                  />
                  {isUploadingCSV && (
                    <div className="absolute inset-0 bg-white/80 flex items-center justify-center">
                      <div className="w-5 h-5 border-2 border-indigo-600 border-t-transparent rounded-full animate-spin" />
                    </div>
                  )}
                </div>
              </div>

              {/* Delete All Data Card */}
              <div className="grid grid-cols-1 gap-4">
                <div 
                  onClick={() => setShowDeleteAllAuth(true)}
                  className="bg-white p-4 rounded-2xl border border-red-100 shadow-sm flex items-center gap-3 relative group overflow-hidden cursor-pointer hover:bg-red-50 transition-colors"
                >
                  <div className="w-10 h-10 bg-red-50 text-red-600 rounded-xl flex items-center justify-center group-hover:bg-red-100 transition-colors">
                    <Trash2 className="w-5 h-5" />
                  </div>
                  <div>
                    <p className="text-[10px] text-red-400 font-bold uppercase tracking-wider">全データ初期化</p>
                    <p className="text-[10px] font-bold text-red-600">システム上のすべてのデータを削除します</p>
                  </div>
                </div>
              </div>

              {/* Settings Section */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div className="bg-white px-5 py-4 rounded-2xl shadow-sm border border-slate-100 flex items-center justify-between">
                  <div className="flex items-center gap-3">
                    <div className="w-12 h-8 bg-slate-50 rounded-lg border border-slate-200 overflow-hidden relative group">
                      {templateUrl ? (
                      <img src={templateUrl} className="w-full h-full object-cover" alt="Template" />
                    ) : (
                      <div className="w-full h-full flex items-center justify-center">
                        <Plus className="w-3 h-3 text-slate-300" />
                      </div>
                    )}
                  </div>
                  <div>
                    <h3 className="text-xs font-bold text-slate-700">パワポ下地設定</h3>
                    <p className="text-[9px] text-slate-400 font-medium">背景画像を適用する</p>
                  </div>
                </div>
                
                <label className="px-3 py-1.5 bg-slate-100 hover:bg-blue-600 hover:text-white text-slate-600 rounded-lg text-[10px] font-bold transition-all cursor-pointer">
                  {templateUrl ? "変更する" : "アップロード"}
                  <input 
                    type="file" 
                    accept="image/*" 
                    className="hidden" 
                    onChange={async (e) => {
                      const file = e.target.files?.[0];
                      if (file) {
                        setIsUploadingTemplate(true);
                        try {
                          const { error } = await supabase.storage
                            .from("reports")
                            .upload("assets/template.jpg", file, { upsert: true });
                          if (error) throw error;
                          
                          const { data: { publicUrl } } = supabase.storage
                            .from("reports")
                            .getPublicUrl("assets/template.jpg");
                          
                          setTemplateUrl(`${publicUrl}?t=${Date.now()}`);
                        } catch (err) {
                          alert("失敗しました");
                        } finally {
                          setIsUploadingTemplate(false);
                        }
                      }
                    }}
                  />
                </label>
                </div>

                <div className="bg-white px-5 py-4 rounded-2xl shadow-sm border border-slate-100 flex items-center justify-between">
                  <div className="flex items-center gap-3">
                    <div className="w-10 h-6 bg-slate-50 border border-slate-200 rounded-lg flex items-center justify-center overflow-hidden flex-shrink-0">
                      <LayoutDashboard className="w-4 h-4 text-slate-400" />
                    </div>
                    <div>
                      <h3 className="text-xs font-bold text-slate-700">パワポ詳細設定</h3>
                      <p className="text-[9px] text-slate-400 font-medium">配置やサイズを詳細調整</p>
                    </div>
                  </div>
                  <button 
                    onClick={() => setShowLayoutSettings(true)}
                    className="px-3 py-1.5 bg-slate-100 hover:bg-blue-600 hover:text-white text-slate-600 rounded-lg text-[10px] font-bold transition-all"
                  >
                    設定する
                  </button>
                </div>
              </div>

              <div className="space-y-3">
                {filteredSubmissions.length === 0 ? (
                  <div className="bg-white p-12 rounded-3xl text-center space-y-4 border border-slate-100 shadow-sm">
                    <div className="w-16 h-16 bg-slate-50 rounded-full flex items-center justify-center mx-auto">
                      <FileText className="w-8 h-8 text-slate-200" />
                    </div>
                    <p className="text-slate-400 text-sm">該当する報告が見つかりません</p>
                  </div>
                ) : (
                  filteredSubmissions.map((sub, i) => (
                    <motion.div 
                      key={i}
                      initial={{ opacity: 0, x: -20 }}
                      animate={{ opacity: 1, x: 0 }}
                      transition={{ delay: i * 0.1 }}
                      className={cn(
                        "bg-white p-5 rounded-2xl shadow-sm border transition-all flex items-center justify-between group",
                        selectedIds.has((sub as any).id) ? "border-blue-500 ring-1 ring-blue-500/20 bg-blue-50/30" : "border-slate-100"
                      )}
                    >
                      <div className="flex items-center gap-4 min-w-0">
                        <div 
                          onClick={() => {
                            const id = (sub as any).id;
                            const newSet = new Set(selectedIds);
                            if (newSet.has(id)) newSet.delete(id);
                            else newSet.add(id);
                            setSelectedIds(newSet);
                          }}
                          className={cn(
                            "w-5 h-5 rounded-md border-2 flex items-center justify-center cursor-pointer transition-all",
                            selectedIds.has((sub as any).id) 
                              ? "bg-blue-600 border-blue-600 text-white" 
                              : "border-slate-200 bg-white"
                          )}
                        >
                          {selectedIds.has((sub as any).id) && <CheckCircle2 className="w-3 h-3" />}
                        </div>
                        <div className="flex-1 min-w-0 space-y-1">
                          {editingSubId === (sub as any).id ? (
                            <div className="space-y-2 p-2 bg-slate-50 rounded-xl border border-slate-200">
                              <input 
                                className="w-full px-2 py-1 text-sm font-bold rounded border" 
                                value={editData.store_name} 
                                onChange={e => setEditData({...editData, store_name: e.target.value})}
                              />
                              <div className="grid grid-cols-2 gap-2">
                                <input 
                                  type="date"
                                  className="px-2 py-1 text-xs rounded border" 
                                  value={editData.completion_date} 
                                  onChange={e => setEditData({...editData, completion_date: e.target.value})}
                                />
                                <input 
                                  className="px-2 py-1 text-xs rounded border" 
                                  value={editData.type} 
                                  onChange={e => setEditData({...editData, type: e.target.value})}
                                />
                              </div>
                              <div className="flex justify-end gap-2 pt-1">
                                <button onClick={() => setEditingSubId(null)} className="text-[10px] font-bold text-slate-400">キャンセル</button>
                                <button onClick={handleEditSave} className="text-[10px] font-bold text-blue-600">保存する</button>
                              </div>
                            </div>
                          ) : (
                            <div 
                              onClick={() => handleEditStart(sub)}
                              className="cursor-pointer group/item min-w-0"
                            >
                              <div className="flex items-start gap-2 min-h-[48px]">
                                <h3 className={cn(
                                  "font-bold transition-all group-hover/item:text-blue-600 line-clamp-2 leading-snug pr-2",
                                  !sub.images ? "text-slate-400 italic" : "text-slate-900"
                                )}>
                                  {(sub as any).store_name || sub.storeName}
                                </h3>
                              </div>
                              <div className="flex items-center gap-3 text-xs text-slate-500 font-semibold truncate mt-1">
                                <span className="shrink-0">{(sub as any).completion_date || sub.completionDate}</span>
                                <span className="shrink-0">•</span>
                                <span className="truncate">{sub.type}</span>
                              </div>
                            </div>
                          )}
                        </div>
                      </div>
                      <div className="flex shrink-0 gap-2 ml-3">
                        {sub.images ? (
                          <>
                            <button 
                              onClick={() => handleDownload(sub, (sub as any).id)}
                              disabled={downloadingId === (sub as any).id}
                              className="p-2.5 text-blue-600 hover:bg-blue-50 rounded-xl transition-colors disabled:opacity-50"
                              title="ダウンロード"
                            >
                              {downloadingId === (sub as any).id ? (
                                <div className="w-5 h-5 border-2 border-blue-600 border-t-transparent rounded-full animate-spin" />
                              ) : (
                                <Download className="w-5 h-5" />
                              )}
                            </button>
                            <button 
                              onClick={() => handleDelete((sub as any).id)}
                              disabled={deletingId === (sub as any).id}
                              className="p-2.5 text-red-400 hover:bg-red-50 rounded-xl transition-colors disabled:opacity-50"
                              title="削除"
                            >
                              {deletingId === (sub as any).id ? (
                                <div className="w-5 h-5 border-2 border-red-400 border-t-transparent rounded-full animate-spin" />
                              ) : (
                                <Trash2 className="w-5 h-5" />
                              )}
                            </button>
                          </>
                        ) : (
                          <div className="flex flex-col items-center gap-1.5">
                            <span className="shrink-0 text-[9px] bg-slate-100 text-slate-500 px-2.5 py-0.5 rounded-full font-bold">報告待ち</span>
                            <div className="flex items-center gap-1.5 px-3 py-1 text-[10px] font-bold text-slate-400 border border-slate-100 rounded-lg whitespace-nowrap">
                              <Clock className="w-3 h-3 shrink-0" />
                              未提出
                            </div>
                          </div>
                        )}
                      </div>
                    </motion.div>
                  ))
                )}
              </div>
            </motion.div>
          )}
        </AnimatePresence>
      </main>

      {/* PPT Layout Settings Modal */}
      <AnimatePresence>
        {showLayoutSettings && (
          <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-sm z-50 flex justify-center items-center p-4">
            <motion.div 
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              className="bg-white rounded-3xl w-full max-w-xl max-h-[90vh] overflow-y-auto shadow-2xl p-6"
            >
              <h2 className="text-lg font-bold text-slate-800 mb-1">パワポレイアウト設定</h2>
              <p className="text-xs text-slate-500 mb-6">出力されるPowerPointの各要素の座標・サイズ（インチ単位）をカスタマイズできます。</p>
              
              <div className="space-y-6">
                {/* 1. Table Settings */}
                <div className="space-y-3">
                  <h3 className="text-sm font-bold text-slate-700 border-b pb-1">表（Table）の配置・サイズ</h3>
                  <div className="grid grid-cols-3 gap-3">
                    <div>
                      <label className="text-[10px] font-bold text-slate-500">X座標 (左から)</label>
                      <input type="number" step="0.1" value={pptLayout.tableX} onChange={e => setPptLayout({...pptLayout, tableX: parseFloat(e.target.value)})} className="w-full px-2 py-1.5 text-sm border rounded bg-slate-50" />
                    </div>
                    <div>
                      <label className="text-[10px] font-bold text-slate-500">Y座標 (上から)</label>
                      <input type="number" step="0.1" value={pptLayout.tableY} onChange={e => setPptLayout({...pptLayout, tableY: parseFloat(e.target.value)})} className="w-full px-2 py-1.5 text-sm border rounded bg-slate-50" />
                    </div>
                    <div>
                      <label className="text-[10px] font-bold text-slate-500">行の高さ</label>
                      <input type="number" step="0.01" value={pptLayout.tableRowH} onChange={e => setPptLayout({...pptLayout, tableRowH: parseFloat(e.target.value)})} className="w-full px-2 py-1.5 text-sm border rounded bg-slate-50" />
                    </div>
                    <div>
                      <label className="text-[10px] font-bold text-slate-500">項目列の幅</label>
                      <input type="number" step="0.1" value={pptLayout.tableW1} onChange={e => setPptLayout({...pptLayout, tableW1: parseFloat(e.target.value)})} className="w-full px-2 py-1.5 text-sm border rounded bg-slate-50" />
                    </div>
                    <div>
                      <label className="text-[10px] font-bold text-slate-500">値列の幅</label>
                      <input type="number" step="0.1" value={pptLayout.tableW2} onChange={e => setPptLayout({...pptLayout, tableW2: parseFloat(e.target.value)})} className="w-full px-2 py-1.5 text-sm border rounded bg-slate-50" />
                    </div>
                    <div>
                      <label className="text-[10px] font-bold text-slate-500">文字サイズ</label>
                      <input type="number" step="1" value={pptLayout.tableFontSize} onChange={e => setPptLayout({...pptLayout, tableFontSize: parseFloat(e.target.value)})} className="w-full px-2 py-1.5 text-sm border rounded bg-slate-50" />
                    </div>
                  </div>
                </div>

                {/* 2. Before/After */}
                <div className="space-y-3">
                  <h3 className="text-sm font-bold text-slate-700 border-b pb-1">比較（BEFORE / AFTER）</h3>
                  <div className="grid grid-cols-4 gap-2">
                    <div className="col-span-4 text-xs font-bold text-slate-600">BEFORE写真</div>
                    <label className="text-[10px] font-bold text-slate-500">X: <input type="number" step="0.1" value={pptLayout.beforeX} onChange={e => setPptLayout({...pptLayout, beforeX: parseFloat(e.target.value)})} className="w-full border rounded px-1" /></label>
                    <label className="text-[10px] font-bold text-slate-500">Y: <input type="number" step="0.1" value={pptLayout.beforeY} onChange={e => setPptLayout({...pptLayout, beforeY: parseFloat(e.target.value)})} className="w-full border rounded px-1" /></label>
                    <label className="text-[10px] font-bold text-slate-500">幅: <input type="number" step="0.1" value={pptLayout.beforeW} onChange={e => setPptLayout({...pptLayout, beforeW: parseFloat(e.target.value)})} className="w-full border rounded px-1" /></label>
                    <label className="text-[10px] font-bold text-slate-500">高さ: <input type="number" step="0.1" value={pptLayout.beforeH} onChange={e => setPptLayout({...pptLayout, beforeH: parseFloat(e.target.value)})} className="w-full border rounded px-1" /></label>
                    
                    <div className="col-span-4 text-xs font-bold text-slate-600 mt-2">AFTER写真</div>
                    <label className="text-[10px] font-bold text-slate-500">X: <input type="number" step="0.1" value={pptLayout.afterX} onChange={e => setPptLayout({...pptLayout, afterX: parseFloat(e.target.value)})} className="w-full border rounded px-1" /></label>
                    <label className="text-[10px] font-bold text-slate-500">Y: <input type="number" step="0.1" value={pptLayout.afterY} onChange={e => setPptLayout({...pptLayout, afterY: parseFloat(e.target.value)})} className="w-full border rounded px-1" /></label>
                    <label className="text-[10px] font-bold text-slate-500">幅: <input type="number" step="0.1" value={pptLayout.afterW} onChange={e => setPptLayout({...pptLayout, afterW: parseFloat(e.target.value)})} className="w-full border rounded px-1" /></label>
                    <label className="text-[10px] font-bold text-slate-500">高さ: <input type="number" step="0.1" value={pptLayout.afterH} onChange={e => setPptLayout({...pptLayout, afterH: parseFloat(e.target.value)})} className="w-full border rounded px-1" /></label>
                  </div>
                </div>

                {/* 3. 3-Way Views */}
                <div className="space-y-3">
                  <h3 className="text-sm font-bold text-slate-700 border-b pb-1">三面写真（正面 / 左側面 / 右側面）</h3>
                  <div className="grid grid-cols-4 gap-2">
                    <div className="col-span-4 text-xs font-bold text-slate-600">正面写真</div>
                    <label className="text-[10px] font-bold text-slate-500">X: <input type="number" step="0.1" value={pptLayout.frontX} onChange={e => setPptLayout({...pptLayout, frontX: parseFloat(e.target.value)})} className="w-full border rounded px-1" /></label>
                    <label className="text-[10px] font-bold text-slate-500">Y: <input type="number" step="0.1" value={pptLayout.frontY} onChange={e => setPptLayout({...pptLayout, frontY: parseFloat(e.target.value)})} className="w-full border rounded px-1" /></label>
                    <label className="text-[10px] font-bold text-slate-500">幅: <input type="number" step="0.1" value={pptLayout.frontW} onChange={e => setPptLayout({...pptLayout, frontW: parseFloat(e.target.value)})} className="w-full border rounded px-1" /></label>
                    <label className="text-[10px] font-bold text-slate-500">高さ: <input type="number" step="0.1" value={pptLayout.frontH} onChange={e => setPptLayout({...pptLayout, frontH: parseFloat(e.target.value)})} className="w-full border rounded px-1" /></label>
                    
                    <div className="col-span-4 text-xs font-bold text-slate-600 mt-2">側面 左</div>
                    <label className="text-[10px] font-bold text-slate-500">X: <input type="number" step="0.1" value={pptLayout.sideLeftX} onChange={e => setPptLayout({...pptLayout, sideLeftX: parseFloat(e.target.value)})} className="w-full border rounded px-1" /></label>
                    <label className="text-[10px] font-bold text-slate-500">Y: <input type="number" step="0.1" value={pptLayout.sideLeftY} onChange={e => setPptLayout({...pptLayout, sideLeftY: parseFloat(e.target.value)})} className="w-full border rounded px-1" /></label>
                    <label className="text-[10px] font-bold text-slate-500">幅: <input type="number" step="0.1" value={pptLayout.sideLeftW} onChange={e => setPptLayout({...pptLayout, sideLeftW: parseFloat(e.target.value)})} className="w-full border rounded px-1" /></label>
                    <label className="text-[10px] font-bold text-slate-500">高さ: <input type="number" step="0.1" value={pptLayout.sideLeftH} onChange={e => setPptLayout({...pptLayout, sideLeftH: parseFloat(e.target.value)})} className="w-full border rounded px-1" /></label>

                    <div className="col-span-4 text-xs font-bold text-slate-600 mt-2">側面 右</div>
                    <label className="text-[10px] font-bold text-slate-500">X: <input type="number" step="0.1" value={pptLayout.sideRightX} onChange={e => setPptLayout({...pptLayout, sideRightX: parseFloat(e.target.value)})} className="w-full border rounded px-1" /></label>
                    <label className="text-[10px] font-bold text-slate-500">Y: <input type="number" step="0.1" value={pptLayout.sideRightY} onChange={e => setPptLayout({...pptLayout, sideRightY: parseFloat(e.target.value)})} className="w-full border rounded px-1" /></label>
                    <label className="text-[10px] font-bold text-slate-500">幅: <input type="number" step="0.1" value={pptLayout.sideRightW} onChange={e => setPptLayout({...pptLayout, sideRightW: parseFloat(e.target.value)})} className="w-full border rounded px-1" /></label>
                    <label className="text-[10px] font-bold text-slate-500">高さ: <input type="number" step="0.1" value={pptLayout.sideRightH} onChange={e => setPptLayout({...pptLayout, sideRightH: parseFloat(e.target.value)})} className="w-full border rounded px-1" /></label>
                  </div>
                </div>

                {/* 4. Other Brands */}
                <div className="space-y-3">
                  <h3 className="text-sm font-bold text-slate-700 border-b pb-1">他社比較（他社1 / 他社2）</h3>
                  <div className="grid grid-cols-4 gap-2">
                    <div className="col-span-4 text-xs font-bold text-slate-600">他社比較 1</div>
                    <label className="text-[10px] font-bold text-slate-500">X: <input type="number" step="0.1" value={pptLayout.other1X} onChange={e => setPptLayout({...pptLayout, other1X: parseFloat(e.target.value)})} className="w-full border rounded px-1" /></label>
                    <label className="text-[10px] font-bold text-slate-500">Y: <input type="number" step="0.1" value={pptLayout.other1Y} onChange={e => setPptLayout({...pptLayout, other1Y: parseFloat(e.target.value)})} className="w-full border rounded px-1" /></label>
                    <label className="text-[10px] font-bold text-slate-500">幅: <input type="number" step="0.1" value={pptLayout.other1W} onChange={e => setPptLayout({...pptLayout, other1W: parseFloat(e.target.value)})} className="w-full border rounded px-1" /></label>
                    <label className="text-[10px] font-bold text-slate-500">高さ: <input type="number" step="0.1" value={pptLayout.other1H} onChange={e => setPptLayout({...pptLayout, other1H: parseFloat(e.target.value)})} className="w-full border rounded px-1" /></label>
                    
                    <div className="col-span-4 text-xs font-bold text-slate-600 mt-2">他社比較 2</div>
                    <label className="text-[10px] font-bold text-slate-500">X: <input type="number" step="0.1" value={pptLayout.other2X} onChange={e => setPptLayout({...pptLayout, other2X: parseFloat(e.target.value)})} className="w-full border rounded px-1" /></label>
                    <label className="text-[10px] font-bold text-slate-500">Y: <input type="number" step="0.1" value={pptLayout.other2Y} onChange={e => setPptLayout({...pptLayout, other2Y: parseFloat(e.target.value)})} className="w-full border rounded px-1" /></label>
                    <label className="text-[10px] font-bold text-slate-500">幅: <input type="number" step="0.1" value={pptLayout.other2W} onChange={e => setPptLayout({...pptLayout, other2W: parseFloat(e.target.value)})} className="w-full border rounded px-1" /></label>
                    <label className="text-[10px] font-bold text-slate-500">高さ: <input type="number" step="0.1" value={pptLayout.other2H} onChange={e => setPptLayout({...pptLayout, other2H: parseFloat(e.target.value)})} className="w-full border rounded px-1" /></label>
                  </div>
                </div>

              </div>
              
              <div className="flex gap-3 mt-8">
                <button 
                  onClick={() => {
                    const confirm = window.confirm("初期設定に戻しますか？");
                    if (confirm) handleSavePptLayout(DEFAULT_LAYOUT);
                  }} 
                  className="flex-1 px-4 py-3 bg-slate-100 text-slate-600 hover:bg-red-50 hover:text-red-600 transition-colors rounded-xl font-bold"
                >
                  初期設定に戻す
                </button>
                <button 
                  onClick={() => handleSavePptLayout(pptLayout)} 
                  className="flex-1 px-4 py-3 bg-blue-600 text-white rounded-xl font-bold shadow-lg shadow-blue-200"
                >
                  保存して閉じる
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Success Notification */}
      <AnimatePresence>
        {showSuccess && (
          <motion.div 
            initial={{ opacity: 0, y: 50 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: 50 }}
            className="fixed bottom-24 left-1/2 -translate-x-1/2 bg-slate-900 text-white px-6 py-3 rounded-full flex items-center gap-3 shadow-2xl z-50 whitespace-nowrap"
          >
            <CheckCircle2 className="w-5 h-5 text-green-400" />
            <span className="text-sm font-medium">報告を送信しました！</span>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Tab Bar (Mobile Navigation Style) - Only for Admin to switch between potentially more views, or just remove if single view */}
      {role === "admin" && (
        <nav className="fixed bottom-0 left-0 right-0 glass shadow-[0_-2px_10px_rgba(0,0,0,0.05)] px-8 py-4 flex justify-around items-center z-50">
          <div className="flex flex-col items-center gap-1 text-blue-600 scale-110">
            <LayoutDashboard className="w-6 h-6" />
            <span className="text-[10px] font-bold">報告一覧</span>
          </div>
        </nav>
      )}
      
      {role === "vendor" && (
        <nav className="fixed bottom-0 left-0 right-0 glass shadow-[0_-2px_10px_rgba(0,0,0,0.05)] px-8 py-4 flex justify-around items-center z-50">
          <div className="flex flex-col items-center gap-1 text-blue-600 scale-110">
            <Camera className="w-6 h-6" />
            <span className="text-[10px] font-bold">報告作成</span>
          </div>
        </nav>
      )}
    </div>
  );
}
