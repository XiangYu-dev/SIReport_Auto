import React from 'react';
import { Outlet, useLocation, useNavigate } from 'react-router-dom';
import { FileSpreadsheet, FileText } from 'lucide-react';

export default function Layout({ onGenerate }) {
  const location = useLocation();
  const navigate = useNavigate();
  
  const menuItems = [
    { path: '/si', label: 'SI文件', icon: FileSpreadsheet },
    { path: '/invoice', label: 'Invoice文件', icon: FileText },
  ];

  const currentPath = location.pathname;

  return (
    <div className="min-h-screen bg-gray-50">
      {/* Header */}
      <div className="sticky top-0 z-50 bg-gradient-to-r from-blue-600 to-blue-700 p-4 md:p-6 shadow-lg">
        <div className="flex items-center justify-between gap-4 max-w-[1600px] mx-auto">
          <div className="text-center md:text-left">
            <h1 className="text-2xl md:text-3xl font-bold text-white">文件自動化系統 v1.0</h1>
            <p className="text-blue-100 mt-1 text-sm md:text-base"></p>
          </div>
          <button 
            onClick={onGenerate}
            className="px-6 py-3 rounded-xl font-bold flex items-center gap-2 bg-white text-blue-600 hover:bg-blue-50 shadow-lg text-sm md:text-base"
          >
            <FileSpreadsheet size={20} />
            產製文件
          </button>
        </div>
      </div>

      <div className="flex max-w-[1600px] mx-auto">
        {/* Sidebar */}
        <aside className="w-48 md:w-56 bg-white border-r border-gray-200 min-h-[calc(100vh-80px)] p-4">
          <nav className="space-y-2">
            {menuItems.map((item) => {
              const Icon = item.icon;
              const isActive = currentPath === item.path || (currentPath === '/' && item.path === '/si');
              return (
                <button
                  key={item.path}
                  onClick={() => navigate(item.path)}
                  className={`w-full flex items-center gap-3 px-4 py-3 rounded-lg text-left transition-all ${
                    isActive 
                      ? 'bg-blue-50 text-blue-700 border-l-4 border-blue-600' 
                      : 'text-gray-600 hover:bg-gray-50 hover:text-gray-800'
                  }`}
                >
                  <Icon size={20} />
                  <span className="font-medium text-sm">{item.label}</span>
                </button>
              );
            })}
          </nav>
        </aside>

        {/* Main Content */}
        <main className="flex-1 p-4 md:p-6">
          <Outlet />
        </main>
      </div>
    </div>
  );
}