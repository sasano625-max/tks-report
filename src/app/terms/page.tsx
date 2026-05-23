import Link from "next/link";
import { ArrowLeft } from "lucide-react";

export default function TermsPage() {
  return (
    <div className="min-h-screen bg-slate-50 p-6 sm:p-12">
      <div className="max-w-3xl mx-auto bg-white p-8 sm:p-12 rounded-3xl shadow-sm border border-slate-100">
        <Link href="/" className="inline-flex items-center text-sm font-bold text-indigo-600 hover:text-indigo-800 mb-8 transition-colors">
          <ArrowLeft className="w-4 h-4 mr-2" /> ホームに戻る
        </Link>
        
        <h1 className="text-3xl font-black text-slate-900 mb-8 tracking-tight">利用規約</h1>
        
        <div className="space-y-8 text-slate-600 leading-relaxed">
          <section>
            <h2 className="text-xl font-bold text-slate-800 mb-4">第1条（適用）</h2>
            <p>本規約は、ユーザーとシステム管理者との間の本サービスの利用に関わる一切の関係に適用されるものとします。</p>
          </section>
          <section>
            <h2 className="text-xl font-bold text-slate-800 mb-4">第2条（写真撮影に関する注意・禁止事項）</h2>
            <p>ユーザーは、本サービスを利用して写真をアップロードする際、以下の行為に十分注意し、これを行ってはなりません。</p>
            <ul className="list-disc list-inside mt-2 space-y-2">
              <li>一般の来店客や通行人の顔など、第三者のプライバシーに関わる被写体を無断で撮影・アップロードする行為</li>
              <li>ポスター、キャラクター、他社の著作物などを、報告の目的を超えて不必要に撮影・アップロードする行為</li>
              <li>オープン前の店舗やバックヤードなど、機密性が高い場所の写真を第三者に漏洩する行為</li>
            </ul>
          </section>
          <section>
            <h2 className="text-xl font-bold text-slate-800 mb-4">第3条（コンテンツの管理）</h2>
            <p>管理者は、不適切と判断した画像データ等を、ユーザーに事前に通知することなく削除できるものとします。</p>
          </section>
          <section>
            <h2 className="text-xl font-bold text-slate-800 mb-4">第4条（免責事項）</h2>
            <p>管理者は、本サービスの利用により発生したユーザーまたは第三者の損害について、一切の責任を負わないものとします。</p>
          </section>
        </div>
      </div>
    </div>
  );
}
