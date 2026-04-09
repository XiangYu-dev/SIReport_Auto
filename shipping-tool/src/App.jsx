import React, { useState } from 'react';
import { FileSpreadsheet, FileText } from 'lucide-react';
import SI from './Page/SI';
import Invoice from './Page/Invoice';

export default function App() {
  const [currentPage, setCurrentPage] = useState('si');

  return (
    <div className="min-h-screen bg-gray-50">
      <div className="sticky top-0 z-50 bg-gradient-to-r from-blue-600 to-blue-700 p-4 md:p-6 shadow-lg">
        <div className="max-w-[1600px] mx-auto">
          <h1 className="text-2xl md:text-3xl font-bold text-white">文件自動化系統</h1>
          <p className="text-blue-100 mt-1 text-sm md:text-base">V1.0</p>
        </div>
      </div>

      <div className="flex max-w-[1600px] mx-auto">
        {/* Sidebar - Tab 切換 */}
        <aside className="sticky top-[80px] w-48 md:w-56 bg-white border-r border-gray-200 h-[calc(100vh-80px)] p-4 shrink-0">
          <nav className="space-y-2">
            <button 
              onClick={() => setCurrentPage('si')}
              className={`w-full flex items-center gap-3 px-4 py-3 rounded-lg text-left transition-all ${
                currentPage === 'si' 
                  ? 'text-blue-700 bg-blue-50 border-l-4 border-blue-600' 
                  : 'text-gray-600 hover:bg-gray-50 hover:text-gray-800'
              }`}
            >
              <FileSpreadsheet size={20} />
              <span className="font-medium text-sm">SI文件</span>
            </button>
            <button 
              onClick={() => setCurrentPage('invoice')}
              className={`w-full flex items-center gap-3 px-4 py-3 rounded-lg text-left transition-all ${
                currentPage === 'invoice' 
                  ? 'text-blue-700 bg-blue-50 border-l-4 border-blue-600' 
                  : 'text-gray-600 hover:bg-gray-50 hover:text-gray-800'
              }`}
            >
              <FileText size={20} />
              <span className="font-medium text-sm">Invoice文件</span>
            </button>
          </nav>
        </aside>

        {/* Main - Tab 內容 */}
        <main className="flex-1 h-[calc(100vh-80px)] overflow-y-auto">
          {currentPage === 'si' && <SI />}
          {currentPage === 'invoice' && <Invoice />}
        </main>
      </div>
    </div>
  );
}