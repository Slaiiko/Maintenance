import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { ProcessedSite, ViewMode, SiteListFilters } from './types';
import FileUpload from './components/FileUpload';
import Dashboard from './components/Dashboard';
import SiteList from './components/SiteList';
import PlanningBoard from './components/PlanningBoard';
import { LayoutDashboard, List, Zap, Menu, X, CalendarCheck, Download, Save, MonitorDown } from 'lucide-react';

const App: React.FC = () => {
  const [data, setData] = useState<ProcessedSite[]>([]);
  const [currentView, setCurrentView] = useState<ViewMode>('dashboard');
  const [sidebarOpen, setSidebarOpen] = useState(true);
  const [siteListFilters, setSiteListFilters] = useState<SiteListFilters>({});
  const [deferredPrompt, setDeferredPrompt] = useState<any>(null);
  const [showInstallBtn, setShowInstallBtn] = useState(false);

  useEffect(() => {
    const handler = (e: any) => {
      e.preventDefault();
      setDeferredPrompt(e);
      setShowInstallBtn(true);
    };
    window.addEventListener('beforeinstallprompt', handler);
    return () => window.removeEventListener('beforeinstallprompt', handler);
  }, []);

  const handleInstallClick = async () => {
    if (!deferredPrompt) return;
    deferredPrompt.prompt();
    const { outcome } = await deferredPrompt.userChoice;
    if (outcome === 'accepted') {
      setShowInstallBtn(false);
    }
    setDeferredPrompt(null);
  };

  const handleDashboardNavigate = (filters: SiteListFilters) => {
    setSiteListFilters(filters);
    setCurrentView('list');
  };

  const handleNavClick = (view: ViewMode) => {
    if (view === 'list') {
       setSiteListFilters({}); // Reset filters when manually clicking menu
    }
    setCurrentView(view);
  };

  const handleExport = () => {
    if (data.length === 0) return;

    // Préparation des données pour l'export propre
    const exportData = data.map(s => ({
      'ID': s.id,
      'Site': s.siteName,
      'Adresse': s.address,
      'Interlocuteur': s.interlocutor,
      'Bornes': s.terminalsRaw,
      'Marque': s.brand,
      'Statut Construction': s.statusConstruction,
      'Date Fin Chantier': s.dateFinChantier ? s.dateFinChantier.toLocaleDateString() : '',
      'Contrat Maintenance': s.hasMaintenanceContract ? 'OUI' : 'NON',
      'Date Début Contrat': s.dateContrat ? s.dateContrat.toLocaleDateString() : '',
      'Statut Contrat (Calculé)': s.contractStatus,
      'Maintenance An 2': s.maintenance1Status,
      'Date Maint. An 2': s.dateMaintenance1Raw,
      'Maintenance An 3': s.maintenance2Status,
      'Date Maint. An 3': s.dateMaintenance2Raw,
      'Action Requise': s.maintenanceNeededCount > 0 ? 'À PLANIFIER' : 'AUCUNE',
      'Retard': s.isLate ? 'OUI' : 'NON'
    }));

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Suivi Consolidé");
    
    const dateStr = new Date().toISOString().split('T')[0];
    XLSX.writeFile(wb, `Suivi_Bornes_Global_${dateStr}.xlsx`);
  };

  if (data.length === 0) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-gradient-to-br from-slate-100 to-slate-200">
        <div className="w-full max-w-2xl bg-white/80 backdrop-blur-xl p-12 rounded-3xl shadow-2xl border border-white/20">
          <div className="text-center mb-10">
            <div className="inline-flex items-center justify-center w-20 h-20 rounded-2xl bg-gradient-to-tr from-blue-600 to-indigo-600 shadow-lg shadow-blue-500/30 mb-6">
              <Zap size={40} className="text-white" />
            </div>
            <h1 className="text-3xl font-extrabold text-slate-800 tracking-tight">EV Contract Manager</h1>
            <p className="text-slate-500 mt-2 text-lg">Analysez vos contrats et maintenances avec précision</p>
          </div>
          <FileUpload onDataLoaded={setData} />
          
          {showInstallBtn && (
            <div className="mt-8 flex justify-center">
              <button 
                onClick={handleInstallClick}
                className="flex items-center gap-2 px-6 py-3 bg-slate-900 text-white rounded-full font-medium hover:bg-slate-800 transition-all shadow-lg hover:shadow-xl hover:-translate-y-1"
              >
                <MonitorDown size={18} />
                Installer l'application sur le bureau
              </button>
            </div>
          )}
        </div>
      </div>
    );
  }

  const NavItem = ({ view, icon, label }: { view: ViewMode, icon: React.ReactNode, label: string }) => (
    <button
      onClick={() => handleNavClick(view)}
      className={`group flex items-center gap-3 w-full p-3.5 rounded-xl text-sm font-medium transition-all duration-300 ease-in-out relative overflow-hidden ${
        currentView === view 
          ? 'bg-blue-600 text-white shadow-lg shadow-blue-900/20' 
          : 'text-slate-400 hover:bg-slate-800/50 hover:text-white'
      }`}
    >
      <div className="relative z-10 flex items-center gap-3">
        {icon}
        <span className={`transition-opacity duration-300 ${sidebarOpen ? 'opacity-100' : 'opacity-0 w-0'}`}>
          {label}
        </span>
      </div>
      {currentView === view && (
        <div className="absolute right-2 top-1/2 -translate-y-1/2 w-1.5 h-1.5 rounded-full bg-white/50" />
      )}
    </button>
  );

  return (
    <div className="flex h-screen bg-slate-50 font-sans overflow-hidden">
      {/* Modern Sidebar */}
      <aside 
        className={`${sidebarOpen ? 'w-72' : 'w-24'} bg-slate-900 transition-all duration-500 ease-[cubic-bezier(0.25,1,0.5,1)] flex flex-col relative z-30 shadow-2xl`}
      >
        <div className="p-6 flex items-center justify-between">
          <div className={`flex items-center gap-3 transition-opacity duration-300 ${sidebarOpen ? 'opacity-100' : 'opacity-0 hidden'}`}>
            <div className="w-8 h-8 rounded-lg bg-gradient-to-br from-blue-500 to-indigo-600 flex items-center justify-center">
              <Zap size={18} className="text-white" />
            </div>
            <span className="font-bold text-xl text-white tracking-wide">EV Manager</span>
          </div>
          <button 
            onClick={() => setSidebarOpen(!sidebarOpen)} 
            className="p-2 rounded-lg bg-slate-800 text-slate-400 hover:text-white hover:bg-slate-700 transition-colors"
          >
            {sidebarOpen ? <X size={20} /> : <Menu size={20} />}
          </button>
        </div>
        
        <nav className="p-4 space-y-2 flex-1 mt-4">
          <NavItem view="dashboard" icon={<LayoutDashboard size={22} />} label="Tableau de Bord" />
          <NavItem view="planning" icon={<CalendarCheck size={22} />} label="Planification" />
          <NavItem view="list" icon={<List size={22} />} label="Liste Complète" />
        </nav>

        <div className="p-4 space-y-3 border-t border-slate-800">
          {showInstallBtn && (
            <button 
              onClick={handleInstallClick}
              className={`w-full flex items-center gap-2 text-sm text-blue-400 hover:text-blue-300 hover:bg-blue-900/20 p-3 rounded-xl transition-all ${!sidebarOpen && 'justify-center'}`}
              title="Installer l'application"
            >
              <MonitorDown size={18} />
              {sidebarOpen && <span>Installer l'App</span>}
            </button>
          )}

          <button 
            onClick={handleExport}
            className={`w-full flex items-center gap-2 text-sm text-emerald-400 hover:text-emerald-300 hover:bg-emerald-900/20 p-3 rounded-xl transition-all ${!sidebarOpen && 'justify-center'}`}
            title="Exporter les données consolidées en Excel"
          >
             <Save size={18} />
             {sidebarOpen && <span>Sauvegarder / Export</span>}
          </button>

          <button 
            onClick={() => setData([])}
            className={`w-full flex items-center gap-2 text-sm text-red-400 hover:text-red-300 hover:bg-red-900/20 p-3 rounded-xl transition-all ${!sidebarOpen && 'justify-center'}`}
            title="Fermer le fichier actuel"
          >
             <X size={18} />
             {sidebarOpen && <span>Fermer le fichier</span>}
          </button>
        </div>
      </aside>

      {/* Main Content Area */}
      <main className="flex-1 flex flex-col h-full relative overflow-hidden bg-[#F8FAFC]">
        {/* Header Content */}
        <header className="px-8 py-8 flex justify-between items-end bg-white/50 backdrop-blur-sm sticky top-0 z-10 border-b border-slate-200/60">
          <div>
             <h1 className="text-3xl font-extrabold text-slate-800 tracking-tight">
              {currentView === 'dashboard' && 'Tableau de Bord'}
              {currentView === 'list' && 'Liste des Contrats'}
              {currentView === 'planning' && 'Centre de Planification'}
            </h1>
            <div className="flex items-center gap-2 mt-2 text-slate-500 font-medium">
              <span className="w-2 h-2 rounded-full bg-green-500 animate-pulse"></span>
              Données chargées : {data.length} sites actifs
            </div>
          </div>
          <div className="hidden md:block">
            <div className="flex items-center gap-4">
               <button 
                  onClick={handleExport}
                  className="flex items-center gap-2 px-4 py-2 bg-slate-900 text-white rounded-lg font-medium text-sm hover:bg-slate-800 transition-colors shadow-lg shadow-slate-900/20"
                >
                  <Download size={16} />
                  Exporter Excel
               </button>
               <div className="text-right pl-4 border-l border-slate-200">
                  <p className="text-xs font-bold text-slate-400 uppercase tracking-wider">Aujourd'hui</p>
                  <p className="text-sm font-semibold text-slate-700">{new Date().toLocaleDateString()}</p>
               </div>
            </div>
          </div>
        </header>

        {/* Scrollable Content */}
        <div className="flex-1 overflow-auto p-8 scrollbar-hide">
          <div className="max-w-7xl mx-auto space-y-8 animate-fade-in-up">
            {currentView === 'dashboard' && <Dashboard data={data} onNavigate={handleDashboardNavigate} />}
            {currentView === 'list' && <SiteList key={JSON.stringify(siteListFilters)} data={data} initialFilters={siteListFilters} />}
            {currentView === 'planning' && <PlanningBoard data={data} />}
          </div>
        </div>
      </main>
    </div>
  );
};

export default App;