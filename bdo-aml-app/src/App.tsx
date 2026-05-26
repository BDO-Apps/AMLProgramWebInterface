import { useState, useEffect } from 'react';
import { 
  LogOut, Sun, Moon, Database, Shield, ArrowLeft, RefreshCw
} from 'lucide-react';
import { msalService } from './services/MSALService';
import type { M365User, AppRole, AppOffice } from './services/MSALService';
import { sharePointService } from './services/SharePointService';
import type { OnboardingCase } from './services/SharePointService';
import { Dashboard } from './components/Dashboard';
import { OnboardingForm } from './components/OnboardingForm';
import { WorkflowPanel } from './components/WorkflowPanel';

function App() {
  const [currentUser, setCurrentUser] = useState<M365User | null>(null);
  const [cases, setCases] = useState<OnboardingCase[]>([]);
  const [view, setView] = useState<'dashboard' | 'form'>('dashboard');
  const [selectedCaseId, setSelectedCaseId] = useState<string | null>(null);
  const [theme, setTheme] = useState<'light' | 'dark'>('light');
  const [isLoading, setIsLoading] = useState(false);
  const [authInitializing, setAuthInitializing] = useState(true);
  const [authError, setAuthError] = useState<string | null>(null);
  const [isSigningIn, setIsSigningIn] = useState(false);

  // Restore MSAL SSO session and subscribe to auth changes
  useEffect(() => {
    const unsubscribe = msalService.subscribe((user) => {
      setCurrentUser(user);
    });

    msalService
      .initialize()
      .catch((err) => {
        console.error('MSAL initialization failed:', err);
        setAuthError('Unable to restore your Microsoft session. Please sign in again.');
      })
      .finally(() => setAuthInitializing(false));

    return unsubscribe;
  }, []);

  // Fetch Cases from SharePoint API Broker when user or view updates
  const syncCases = async () => {
    if (currentUser) {
      setIsLoading(true);
      try {
        const fetched = await sharePointService.getCases(currentUser.role, currentUser.email);
        setCases(fetched);
      } catch (e) {
        const message =
          e instanceof Error ? e.message : 'Failed to load cases from SharePoint.';
        console.error('SharePoint fetch error:', e);
        alert(
          `SharePoint error: ${message}\n\nEnsure you are signed in, API permissions (Sites.ReadWrite.All, Files.ReadWrite.All) are granted, and list column internal names match docs/SHAREPOINT_SCHEMA.md.`
        );
      } finally {
        setIsLoading(false);
      }
    }
  };

  useEffect(() => {
    syncCases();
  }, [currentUser]);

  // Handle Theme switching
  useEffect(() => {
    document.documentElement.setAttribute('data-theme', theme);
  }, [theme]);

  const toggleTheme = () => {
    setTheme(prev => prev === 'light' ? 'dark' : 'light');
  };

  // --- Demo Control actions ---
  const handleRoleChange = (role: AppRole, office: AppOffice) => {
    msalService.changeRole(role, office);
    setView('dashboard');
    setSelectedCaseId(null);
  };

  const handleSignIn = async () => {
    setAuthError(null);
    setIsSigningIn(true);
    try {
      await msalService.login();
    } catch (err) {
      const message =
        err instanceof Error ? err.message : 'Microsoft sign-in was cancelled or failed.';
      setAuthError(message);
      console.error('MSAL login error:', err);
    } finally {
      setIsSigningIn(false);
    }
  };

  const handleResetDatabase = () => {
    if (window.confirm("Restore the SharePoint Simulation Database to seeded initial cases? This will clear local changes.")) {
      sharePointService.clearStorage();
      syncCases();
      setView('dashboard');
      setSelectedCaseId(null);
      alert("SharePoint and Document Library lists successfully re-seeded!");
    }
  };

  const handleDiagnoseColumns = async () => {
    try {
      const cols = await sharePointService.diagnoseColumns();
      console.log('--- SHAREPOINT CLIENTONBOARDING COLUMNS DIAGNOSTICS ---');
      console.table(cols);
      
      const pepCol = cols.find(
        (c) =>
          c.name.toLowerCase() === 'cddpepscreened' ||
          c.displayName.toLowerCase() === 'pep screened' ||
          c.name.toLowerCase().includes('pep')
      );
      
      let message = `Successfully audited ${cols.length} columns on SharePoint List 'ClientOnboarding'!\n\nOpen Browser DevTools (F12) → Console tab to inspect the full table of column internal names.\n\n`;
      
      if (pepCol) {
        message += `AUDIT FINDING:\nFound field for PEP screening: Internal Name = '${pepCol.name}', Display Name = '${pepCol.displayName}'.\n\n`;
        if (pepCol.name !== 'CDDPEPScreened') {
          message += `🚨 MISMATCH DETECTED!\nYour SharePoint list column internal name is '${pepCol.name}' but the code is looking for 'CDDPEPScreened'.\n\nTo resolve the error, please rename the column in SharePoint to match the internal name 'CDDPEPScreened', OR update the fieldMapper.ts file to use '${pepCol.name}'.`;
        } else {
          message += `✅ MATCH CONFIRMED!\nYour SharePoint column internal name is perfectly set to 'CDDPEPScreened'.`;
        }
      } else {
        message += `🚨 AUDIT WARNING:\nNo Yes/No column representing 'PEP Screened' was identified on your SharePoint List. Please make sure that you have added the 'PEP screened' column to your ClientOnboarding list.`;
      }
      
      alert(message);
    } catch (e) {
      alert(`Diagnostics failed: ${e}\n\nEnsure that you have granted Sites.ReadWrite.All admin consent and that the SharePoint URL is configured correctly.`);
    }
  };

  // --- Navigation routers ---
  const handleOpenCase = (id: string) => {
    setSelectedCaseId(id);
    setView('form');
  };

  const handleCreateNewCase = () => {
    setSelectedCaseId(null);
    setView('form');
  };

  const handleSaveSuccess = (updatedCase: OnboardingCase) => {
    syncCases();
    // Keep form open but refresh details
    setSelectedCaseId(updatedCase.id);
  };

  if (authInitializing) {
    return (
      <div style={{
        display: 'flex',
        flexDirection: 'column',
        alignItems: 'center',
        justifyContent: 'center',
        height: '100vh',
        backgroundColor: '#0b1319',
        color: 'white',
        fontFamily: 'sans-serif',
        gap: '1rem',
      }}>
        <RefreshCw size={32} color="#00a3a3" style={{ animation: 'spin 1.5s linear infinite' }} />
        <p style={{ opacity: 0.8 }}>Checking Microsoft 365 session…</p>
      </div>
    );
  }

  if (!currentUser) {
    return (
      <div style={{ 
        display: 'flex', 
        flexDirection: 'column', 
        alignItems: 'center', 
        justifyContent: 'center', 
        height: '100vh',
        backgroundColor: '#0b1319',
        color: 'white',
        fontFamily: 'sans-serif'
      }}>
        <Shield size={64} color="#00a3a3" style={{ marginBottom: '1.5rem' }} />
        <h2>BDO Zimbabwe / Malawi</h2>
        <p style={{ opacity: 0.7, marginTop: '0.5rem', marginBottom: '1.5rem' }}>M365 Enterprise AML Compliance Portal</p>
        {authError && (
          <p style={{ color: '#f87171', marginBottom: '1rem', maxWidth: '360px', textAlign: 'center', fontSize: '0.9rem' }}>
            {authError}
          </p>
        )}
        <button 
          className="btn btn-primary" 
          onClick={handleSignIn}
          disabled={isSigningIn}
        >
          {isSigningIn ? 'Signing in…' : 'Sign in with Microsoft'}
        </button>
        <p style={{ opacity: 0.5, marginTop: '1.5rem', fontSize: '0.75rem', maxWidth: '400px', textAlign: 'center' }}>
          Sign in with your BDO account. Role is assigned from your email (Preparer, Compliance, Engagement Partner, or Risk Partner).
        </p>
      </div>
    );
  }

  // Active onboarding details loaded in form
  const activeCaseData = selectedCaseId ? cases.find(c => c.id === selectedCaseId) : null;

  return (
    <div className="app-container">
      
      {/* Header bar */}
      <header className="header">
        <div className="logo-section">
          <div className="bdo-logo">
            BDO <span>Zimbabwe/Malawi</span>
          </div>
          <span className="tagline">AML/CFT/CPF RBA</span>
        </div>

        <div style={{ display: 'flex', alignItems: 'center', gap: '1.5rem' }}>
          
          {/* Theme switcher */}
          <button 
            className="btn" 
            style={{ padding: '8px', borderRadius: '50%', border: 'none', backgroundColor: 'var(--bg-app)', color: 'var(--color-text-secondary)' }}
            onClick={toggleTheme}
          >
            {theme === 'light' ? <Moon size={18} /> : <Sun size={18} />}
          </button>

          {/* User profile dropdown info */}
          <div style={{ display: 'flex', alignItems: 'center', gap: '0.75rem' }}>
            <div style={{ textAlign: 'right', display: 'flex', flexDirection: 'column' }}>
              <span style={{ fontWeight: 700, fontSize: '0.85rem' }}>{currentUser.name}</span>
              <span style={{ fontSize: '0.7rem', color: 'var(--color-primary)', fontWeight: 600 }}>
                {currentUser.role} &bull; {currentUser.office}
              </span>
            </div>
            
            {currentUser.avatar ? (
              <img 
                src={currentUser.avatar} 
                alt="Avatar" 
                style={{ width: '38px', height: '38px', borderRadius: '50%', border: '2px solid var(--color-primary)' }}
              />
            ) : (
              <div style={{ 
                width: '38px', 
                height: '38px', 
                borderRadius: '50%', 
                backgroundColor: 'var(--color-primary-light)', 
                color: 'var(--color-primary)',
                fontWeight: 700,
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                fontSize: '0.9rem'
              }}>
                {currentUser.name.charAt(0)}
              </div>
            )}

            <button 
              className="btn btn-secondary" 
              style={{ padding: '6px', minWidth: 'auto', border: 'none' }}
              onClick={() => msalService.logout()}
            >
              <LogOut size={16} />
            </button>
          </div>

        </div>
      </header>

      {/* Main app body */}
      <main className="main-content">
        {isLoading && (
          <div style={{ 
            display: 'flex', 
            alignItems: 'center', 
            justifyContent: 'center', 
            gap: '0.5rem',
            padding: '1rem',
            backgroundColor: 'var(--color-primary-light)',
            borderRadius: '8px',
            color: 'var(--color-primary)',
            fontSize: '0.85rem',
            fontWeight: 600
          }}>
            <RefreshCw size={16} className="animate-spin" style={{ animation: 'spin 1.5s linear infinite' }} />
            Retrieving Active SharePoint Online Relational Matrices...
          </div>
        )}

        {view === 'dashboard' ? (
          <Dashboard 
            cases={cases}
            currentUser={currentUser}
            onSelectCase={handleOpenCase}
            onNewCase={handleCreateNewCase}
            onRefresh={syncCases}
          />
        ) : (
          <div style={{ display: 'flex', flexDirection: 'column', gap: '1.5rem' }}>
            
            {/* Back to dashboard header link */}
            <button 
              className="btn btn-secondary" 
              style={{ width: 'fit-content' }}
              onClick={() => setView('dashboard')}
            >
              <ArrowLeft size={16} />
              Back to Registry
            </button>

            {/* Split review view if case exists and user is compliance / partner reviewer */}
            {activeCaseData && currentUser.role !== 'Preparer' ? (
              <div style={{ display: 'grid', gridTemplateColumns: '1.2fr 1fr', gap: '1.5rem', alignItems: 'start' }}>
                
                {/* Left Side: Onboarding Form (Read-only format) */}
                <div>
                  <OnboardingForm 
                    caseId={selectedCaseId}
                    currentUser={currentUser}
                    onClose={() => setView('dashboard')}
                    onSaveSuccess={handleSaveSuccess}
                  />
                </div>

                {/* Right Side: Compliance Reviews timelines and Action triggers */}
                <div>
                  <WorkflowPanel 
                    activeCase={activeCaseData}
                    currentUser={currentUser}
                    onActionSuccess={handleSaveSuccess}
                  />
                </div>

              </div>
            ) : (
              // Full size form for Preparers (or creation mode)
              <OnboardingForm 
                caseId={selectedCaseId}
                currentUser={currentUser}
                onClose={() => setView('dashboard')}
                onSaveSuccess={handleSaveSuccess}
              />
            )}

          </div>
        )}
      </main>

      {/* Compliance Walkthrough — test accounts only (implementation plan §1) */}
      {currentUser.canOverrideRole && (
        <div className="control-bar">
          <span className="control-bar-title">Compliance Walkthrough</span>

          <div style={{ display: 'flex', alignItems: 'center', gap: '0.35rem', fontSize: '0.75rem', fontWeight: 600 }}>
            <span>Role:</span>
            <select
              className="form-control"
              style={{ padding: '2px 24px 2px 8px', fontSize: '0.725rem', width: 'auto', border: 'none', backgroundColor: 'var(--bg-app)', height: '24px' }}
              value={currentUser.role}
              onChange={(e) => handleRoleChange(e.target.value as AppRole, currentUser.office)}
            >
              <option value="Preparer">Preparer (Engagement)</option>
              <option value="Compliance">Compliance MLRO</option>
              <option value="EngagementPartner">Engagement Partner</option>
              <option value="RiskPartner">Risk Partner (Senior)</option>
            </select>
          </div>

          <div style={{ display: 'flex', alignItems: 'center', gap: '0.35rem', fontSize: '0.75rem', fontWeight: 600 }}>
            <span>Office:</span>
            <select
              className="form-control"
              style={{ padding: '2px 24px 2px 8px', fontSize: '0.725rem', width: 'auto', border: 'none', backgroundColor: 'var(--bg-app)', height: '24px' }}
              value={currentUser.office}
              onChange={(e) => handleRoleChange(currentUser.role, e.target.value as AppOffice)}
            >
              <option value="Zimbabwe">Zimbabwe</option>
              <option value="Malawi">Malawi</option>
            </select>
          </div>

          <button
            className="btn btn-secondary"
            style={{ padding: '2px 8px', fontSize: '0.725rem', borderRadius: '4px', border: 'none', backgroundColor: 'var(--bg-app)', height: '24px', display: 'flex', alignItems: 'center', gap: '0.25rem' }}
            onClick={handleResetDatabase}
            title="Restore simulation state"
          >
            <Database size={12} />
            Reset DB
          </button>

          {!sharePointService.isDemoMode && (
            <button
              className="btn btn-secondary"
              style={{ padding: '2px 8px', fontSize: '0.725rem', borderRadius: '4px', border: 'none', backgroundColor: 'var(--bg-app)', height: '24px', display: 'flex', alignItems: 'center', gap: '0.25rem', color: 'var(--color-primary)' }}
              onClick={handleDiagnoseColumns}
              title="Audit SharePoint List Columns"
            >
              <Database size={12} />
              Diagnose List
            </button>
          )}
        </div>
      )}

    </div>
  );
}

export default App;
