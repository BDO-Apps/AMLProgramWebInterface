import React, { useState, useMemo } from 'react';
import { 
  FileText, CheckCircle, Clock, Search, Plus, 
  ArrowRight, RefreshCw, BarChart2, ShieldAlert, Award
} from 'lucide-react';
import type { OnboardingCase } from '../services/SharePointService';
import type { M365User } from '../services/MSALService';

interface DashboardProps {
  cases: OnboardingCase[];
  currentUser: M365User;
  onSelectCase: (caseId: string) => void;
  onNewCase: () => void;
  onRefresh: () => void;
}

export const Dashboard: React.FC<DashboardProps> = ({
  cases,
  currentUser,
  onSelectCase,
  onNewCase,
  onRefresh
}) => {
  const [searchTerm, setSearchTerm] = useState('');
  const [statusFilter, setStatusFilter] = useState<string>('All');
  const [sortBy, setSortBy] = useState<'name' | 'risk' | 'date'>('date');

  // Compute stats
  const stats = useMemo(() => {
    const total = cases.length;
    const pending = cases.filter(c => c.status.startsWith('Pending')).length;
    const highRisk = cases.filter(c => c.overallRiskRating === 'High').length;
    const approved = cases.filter(c => c.status === 'Approved').length;
    const upcomingReviews = cases.filter(c => c.status === 'Approved' && c.nextReviewDate).length;

    return { total, pending, highRisk, approved, upcomingReviews };
  }, [cases]);

  // Dynamic compliance alerts
  const alerts = useMemo(() => {
    const list: { id: string; type: 'danger' | 'warning' | 'info'; message: string; caseId?: string }[] = [];
    
    cases.forEach(c => {
      if (c.status.startsWith('Pending') && c.overallRiskRating === 'High') {
        list.push({
          id: `alert-high-${c.id}`,
          type: 'danger',
          message: `HIGH-RISK CLIENT: "${c.clientName}" requires immediate MLRO and Risk Partner verification.`,
          caseId: c.id
        });
      }
      if (c.pepStatus.isPep && c.status === 'Pending Compliance') {
        list.push({
          id: `alert-pep-${c.id}`,
          type: 'danger',
          message: `PEP FLAGGED: Natural person inside "${c.clientName}" is marked as PEP. Source of funds mandatory.`,
          caseId: c.id
        });
      }
      if (c.status === 'Returned') {
        list.push({
          id: `alert-ret-${c.id}`,
          type: 'warning',
          message: `RETURNED: Onboarding for "${c.clientName}" returned by reviewer with comments.`,
          caseId: c.id
        });
      }
      // Check document uploads
      if (c.clientType === 'Legal Person' && !c.documents.some(d => d.category === 'CR6_Directors')) {
        list.push({
          id: `alert-cr6-${c.id}`,
          type: 'warning',
          message: `DOC MISSING: "${c.clientName}" (Legal Person) requires CR6 Directors registry.`,
          caseId: c.id
        });
      }
    });

    // Seed default info alert if list is empty
    if (list.length === 0) {
      list.push({
        id: 'alert-info-default',
        type: 'info',
        message: 'All submitted cases currently satisfy core CDD documentation guidelines.'
      });
    }

    return list;
  }, [cases]);

  // Filter & Sort Cases
  const filteredCases = useMemo(() => {
    let result = [...cases];

    // Filter by search
    if (searchTerm.trim() !== '') {
      const q = searchTerm.toLowerCase();
      result = result.filter(c => 
        c.clientName.toLowerCase().includes(q) || 
        c.id.toLowerCase().includes(q) ||
        c.regNumber.toLowerCase().includes(q) ||
        c.natureOfBusiness.toLowerCase().includes(q)
      );
    }

    // Filter by status tab
    if (statusFilter !== 'All') {
      if (statusFilter === 'Pending') {
        result = result.filter(c => c.status.startsWith('Pending'));
      } else {
        result = result.filter(c => c.status === statusFilter);
      }
    }

    // Sorting
    result.sort((a, b) => {
      if (sortBy === 'name') {
        return a.clientName.localeCompare(b.clientName);
      }
      if (sortBy === 'risk') {
        const riskWeight = { High: 3, Medium: 2, Low: 1 };
        return riskWeight[b.overallRiskRating] - riskWeight[a.overallRiskRating];
      }
      // Default sort by date / case ID desc
      return b.id.localeCompare(a.id);
    });

    return result;
  }, [cases, searchTerm, statusFilter, sortBy]);

  // SVG Chart Data - Risk Breakdown
  const riskChartData = useMemo(() => {
    const low = cases.filter(c => c.overallRiskRating === 'Low').length;
    const med = cases.filter(c => c.overallRiskRating === 'Medium').length;
    const high = cases.filter(c => c.overallRiskRating === 'High').length;
    const total = low + med + high || 1;

    return {
      lowPct: (low / total) * 100,
      medPct: (med / total) * 100,
      highPct: (high / total) * 100,
      low, med, high
    };
  }, [cases]);

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: '1.5rem' }}>
      
      {/* Upper Analytics Banner */}
      <div className="flex-between">
        <div>
          <h1 style={{ fontSize: '1.75rem', fontWeight: 700, color: 'var(--color-text-main)' }}>
            AML / CFT / CPF Onboarding Registry
          </h1>
          <p style={{ color: 'var(--color-text-secondary)', fontSize: '0.875rem', marginTop: '0.25rem' }}>
            Welcome back, <strong style={{ color: 'var(--color-primary)' }}>{currentUser.name}</strong> ({currentUser.role}) &bull; BDO {currentUser.office} Office
          </p>
        </div>
        
        <div style={{ display: 'flex', gap: '0.75rem' }}>
          <button className="btn btn-secondary" onClick={onRefresh}>
            <RefreshCw size={16} />
            Sync Data
          </button>
          
          {currentUser.role === 'Preparer' && (
            <button className="btn btn-primary" onClick={onNewCase}>
              <Plus size={16} />
              New Client Form
            </button>
          )}
        </div>
      </div>

      {/* KPI Stats Cards */}
      <div className="stat-card-grid">
        <div className="stat-card">
          <div className="stat-icon-wrapper" style={{ backgroundColor: 'var(--color-pending-bg)', color: 'var(--color-pending)' }}>
            <Clock size={24} />
          </div>
          <div className="stat-info">
            <span className="stat-val">{stats.pending}</span>
            <span className="stat-lbl">Pending Review</span>
          </div>
        </div>

        <div className="stat-card">
          <div className="stat-icon-wrapper" style={{ backgroundColor: 'var(--color-high-risk-bg)', color: 'var(--color-high-risk)' }}>
            <ShieldAlert size={24} />
          </div>
          <div className="stat-info">
            <span className="stat-val">{stats.highRisk}</span>
            <span className="stat-lbl">High Risk Active</span>
          </div>
        </div>

        <div className="stat-card">
          <div className="stat-icon-wrapper" style={{ backgroundColor: 'var(--color-low-risk-bg)', color: 'var(--color-low-risk)' }}>
            <CheckCircle size={24} />
          </div>
          <div className="stat-info">
            <span className="stat-val">{stats.approved}</span>
            <span className="stat-lbl">Approved Clients</span>
          </div>
        </div>

        <div className="stat-card">
          <div className="stat-icon-wrapper" style={{ backgroundColor: 'var(--color-returned-bg)', color: 'var(--color-returned)' }}>
            <Award size={24} />
          </div>
          <div className="stat-info">
            <span className="stat-val">{stats.upcomingReviews}</span>
            <span className="stat-lbl">Scheduled Audits</span>
          </div>
        </div>
      </div>

      {/* Double Column Grid: Charts / Banners & Registry */}
      <div style={{ display: 'grid', gridTemplateColumns: '320px 1fr', gap: '1.5rem' }}>
        
        {/* Left Hand Column: Alerts & Analytics */}
        <div style={{ display: 'flex', flexDirection: 'column', gap: '1.5rem' }}>
          
          {/* Risk Profile Distribution Chart */}
          <div className="card">
            <div className="flex-between" style={{ marginBottom: '1rem' }}>
              <h3 style={{ fontSize: '0.9rem', fontWeight: 700, textTransform: 'uppercase', color: 'var(--color-text-secondary)', letterSpacing: '0.05em' }}>
                Inherent Risk Ratios
              </h3>
              <BarChart2 size={16} style={{ color: 'var(--color-primary)' }} />
            </div>
            
            {/* Visual SVG Donut/Stacked Chart */}
            <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', margin: '1rem 0' }}>
              <svg width="180" height="20" style={{ borderRadius: '4px', overflow: 'hidden', backgroundColor: 'var(--bg-app)' }}>
                {cases.length > 0 ? (
                  <>
                    <rect x="0" y="0" width={`${riskChartData.lowPct}%`} height="20" fill="var(--color-low-risk)" />
                    <rect x={`${riskChartData.lowPct}%`} y="0" width={`${riskChartData.medPct}%`} height="20" fill="var(--color-med-risk)" />
                    <rect x={`${riskChartData.lowPct + riskChartData.medPct}%`} y="0" width={`${riskChartData.highPct}%`} height="20" fill="var(--color-high-risk)" />
                  </>
                ) : (
                  <rect x="0" y="0" width="100%" height="20" fill="var(--color-border)" />
                )}
              </svg>
              
              <div style={{ width: '100%', display: 'flex', justifyContent: 'space-between', fontSize: '0.75rem', marginTop: '0.75rem', fontWeight: 600 }}>
                <span style={{ color: 'var(--color-low-risk)' }}>Low: {riskChartData.low}</span>
                <span style={{ color: 'var(--color-med-risk)' }}>Medium: {riskChartData.med}</span>
                <span style={{ color: 'var(--color-high-risk)' }}>High: {riskChartData.high}</span>
              </div>
            </div>
            
            <p style={{ fontSize: '0.75rem', color: 'var(--color-text-muted)', lineHeight: '1.4' }}>
              BDO Zimbabwe & Malawi enforce enhanced risk mitigation protocols for all entities flagged under Medium or High tiers.
            </p>
          </div>

          {/* Compliance Alerts Panel */}
          <div className="card" style={{ display: 'flex', flexDirection: 'column', gap: '0.75rem' }}>
            <h3 style={{ fontSize: '0.9rem', fontWeight: 700, textTransform: 'uppercase', color: 'var(--color-text-secondary)', letterSpacing: '0.05em', marginBottom: '0.25rem' }}>
              Compliance Flags
            </h3>
            
            <div style={{ display: 'flex', flexDirection: 'column', gap: '0.75rem', maxHeight: '350px', overflowY: 'auto' }}>
              {alerts.map((alert) => (
                <div 
                  key={alert.id} 
                  className={`alert-panel ${alert.type === 'danger' ? 'danger' : ''}`}
                  style={{ 
                    margin: 0, 
                    padding: '0.75rem', 
                    borderRadius: '8px',
                    borderLeftWidth: '3px',
                    fontSize: '0.8rem',
                    cursor: alert.caseId ? 'pointer' : 'default',
                    transition: 'transform var(--transition-fast)'
                  }}
                  onClick={() => alert.caseId && onSelectCase(alert.caseId)}
                >
                  <div style={{ display: 'flex', gap: '0.5rem', alignItems: 'flex-start' }}>
                    <ShieldAlert size={16} style={{ flexShrink: 0, marginTop: '1px', color: alert.type === 'danger' ? 'var(--color-high-risk)' : 'var(--color-primary)' }} />
                    <div>
                      {alert.message}
                      {alert.caseId && (
                        <div style={{ marginTop: '0.25rem', fontWeight: 700, display: 'flex', alignItems: 'center', gap: '0.25rem' }}>
                          View Case <ArrowRight size={10} />
                        </div>
                      )}
                    </div>
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>

        {/* Right Hand Column: Case Registry */}
        <div className="card" style={{ padding: '1.25rem', display: 'flex', flexDirection: 'column', gap: '1.25rem' }}>
          
          {/* Search, Filter, Sort Inputs */}
          <div style={{ display: 'flex', flexWrap: 'wrap', gap: '0.75rem', justifyContent: 'space-between', alignItems: 'center' }}>
            
            {/* Search Input */}
            <div style={{ position: 'relative', flex: 1, minWidth: '240px' }}>
              <Search size={16} style={{ position: 'absolute', left: '0.75rem', top: '50%', transform: 'translateY(-50%)', color: 'var(--color-text-muted)' }} />
              <input 
                type="text" 
                placeholder="Search by client name, reg, industry..." 
                className="form-control"
                style={{ paddingLeft: '2.25rem' }}
                value={searchTerm}
                onChange={e => setSearchTerm(e.target.value)}
              />
            </div>

            {/* Filter and Sort Tabs */}
            <div style={{ display: 'flex', gap: '0.5rem', flexWrap: 'wrap' }}>
              <div style={{ display: 'flex', gap: '2px', backgroundColor: 'var(--bg-app)', padding: '2px', borderRadius: '8px', border: '1px solid var(--color-border)' }}>
                {['All', 'Draft', 'Pending', 'Returned', 'Approved'].map(tab => (
                  <button
                    key={tab}
                    className="btn"
                    style={{ 
                      padding: '0.35rem 0.75rem', 
                      fontSize: '0.75rem', 
                      borderRadius: '6px',
                      backgroundColor: statusFilter === tab ? 'var(--bg-card)' : 'transparent',
                      color: statusFilter === tab ? 'var(--color-primary)' : 'var(--color-text-secondary)',
                      boxShadow: statusFilter === tab ? 'var(--shadow-sm)' : 'none',
                      fontWeight: statusFilter === tab ? 600 : 500
                    }}
                    onClick={() => setStatusFilter(tab)}
                  >
                    {tab}
                  </button>
                ))}
              </div>

              <select 
                className="form-control" 
                style={{ width: 'auto', fontSize: '0.8rem', padding: '0.35rem 1.75rem 0.35rem 0.75rem' }}
                value={sortBy}
                onChange={e => setSortBy(e.target.value as any)}
              >
                <option value="date">Sort: Recent Created</option>
                <option value="risk">Sort: Risk Level</option>
                <option value="name">Sort: Alphabetical</option>
              </select>
            </div>

          </div>

          {/* Core Table */}
          <div className="table-wrapper">
            <table className="bdo-table">
              <thead>
                <tr>
                  <th>Case ID</th>
                  <th>Client / Entity Name</th>
                  <th>Entity Type</th>
                  <th>Office</th>
                  <th>Risk Rating</th>
                  <th>Workflow Status</th>
                  <th>Last Update</th>
                  <th style={{ width: '80px' }}>Action</th>
                </tr>
              </thead>
              <tbody>
                {filteredCases.length > 0 ? (
                  filteredCases.map(c => {
                    // Risk badge class mapping
                    let riskBadge = 'badge-low';
                    if (c.overallRiskRating === 'Medium') riskBadge = 'badge-med';
                    if (c.overallRiskRating === 'High') riskBadge = 'badge-high';

                    // Status badge class mapping
                    let statusBadge = 'badge-draft';
                    if (c.status.startsWith('Pending')) statusBadge = 'badge-pending';
                    if (c.status === 'Returned') statusBadge = 'badge-returned';
                    if (c.status === 'Rejected') statusBadge = 'badge-rejected';
                    if (c.status === 'Approved') statusBadge = 'badge-low';

                    return (
                      <tr key={c.id}>
                        <td style={{ fontWeight: 600, color: 'var(--color-primary)' }}>{c.id}</td>
                        <td>
                          <div style={{ display: 'flex', flexDirection: 'column' }}>
                            <span style={{ fontWeight: 600 }}>{c.clientName}</span>
                            <span style={{ fontSize: '0.75rem', color: 'var(--color-text-secondary)' }}>{c.natureOfBusiness}</span>
                          </div>
                        </td>
                        <td>{c.clientType}</td>
                        <td style={{ fontWeight: 500 }}>{c.office}</td>
                        <td>
                          <span className={`badge ${riskBadge}`}>{c.overallRiskRating}</span>
                        </td>
                        <td>
                          <span className={`badge ${statusBadge}`}>{c.status}</span>
                        </td>
                        <td style={{ color: 'var(--color-text-secondary)', fontSize: '0.8rem' }}>{c.lastUpdated}</td>
                        <td>
                          <button 
                            className="btn btn-secondary"
                            style={{ padding: '0.25rem 0.5rem', fontSize: '0.75rem', borderRadius: '4px' }}
                            onClick={() => onSelectCase(c.id)}
                          >
                            Open
                            <ArrowRight size={12} />
                          </button>
                        </td>
                      </tr>
                    );
                  })
                ) : (
                  <tr>
                    <td colSpan={8} style={{ textAlign: 'center', padding: '3rem', color: 'var(--color-text-muted)' }}>
                      <FileText size={40} style={{ margin: '0 auto 0.75rem', display: 'block', opacity: 0.5 }} />
                      No client onboarding records match current criteria.
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>

        </div>

      </div>

    </div>
  );
};
