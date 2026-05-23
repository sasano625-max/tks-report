import Link from "next/link";
import { ArrowLeft } from "lucide-react";

export default function PrivacyPage() {
  return (
    <div className="min-h-screen bg-slate-50 p-6 sm:p-12">
      <div className="max-w-3xl mx-auto bg-white p-8 sm:p-12 rounded-3xl shadow-sm border border-slate-100">
        <Link href="/" className="inline-flex items-center text-sm font-bold text-indigo-600 hover:text-indigo-800 mb-8 transition-colors">
          <ArrowLeft className="w-4 h-4 mr-2" /> ホームに戻る
        </Link>
        
        <h1 className="text-3xl font-black text-slate-900 mb-8 tracking-tight">プライバシーポリシー</h1>
        
        <div className="space-y-8 text-slate-600 leading-relaxed">
          <section>
            <h2 className="text-xl font-bold text-slate-800 mb-4">1. 収集する情報</h2>
            <p>本サービスでは、写真報告に関連する以下の情報を収集・保存します。</p>
            <ul className="list-disc list-inside mt-2 space-y-2">
              <li>アップロードされた画像データ</li>
              <li>店舗名、担当者名（入力された場合）などの報告メタデータ</li>
            </ul>
          </section>
          <section>
            <h2 className="text-xl font-bold text-slate-800 mb-4">2. 情報の利用目的</h2>
            <p>収集した情報は、現場状況の確認、設置報告書の自動生成、および本サービスの提供・改善の目的でのみ利用します。</p>
          </section>
          <section>
            <h2 className="text-xl font-bold text-slate-800 mb-4">3. 情報の安全管理</h2>
            <p>収集したデータは、クラウド環境（Supabase）にて安全に保管し、アクセス権限を適切に管理することで情報漏洩の防止に努めます。</p>
          </section>
          <section>
            <h2 className="text-xl font-bold text-slate-800 mb-4">4. 第三者提供</h2>
            <p>当社は、法令に基づく場合を除き、ユーザーの同意を得ることなく第三者に情報を提供することはありません。</p>
          </section>
          <section>
            <h2 className="text-xl font-bold text-slate-800 mb-4">5. お問い合わせ</h2>
            <p>本ポリシーに関するお問い合わせやデータの削除要請は、システム管理者までお願いいたします。</p>
          </section>
        </div>
      </div>
    </div>
  );
}
