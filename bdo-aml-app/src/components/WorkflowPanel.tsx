import React, { useState, useMemo } from 'react';
import { 
  CheckCircle, ArrowLeftRight, XCircle, Clock, ShieldCheck, 
  MessageSquare, UserCheck
} from 'lucide-react';
import { sharePointService } from '../services/SharePointService';
import type { OnboardingCase } from '../services/SharePointService';
import type { M365User } from '../services/MSALService';

interface WorkflowPanelProps {
  activeCase: OnboardingCase;
  currentUser: M365User;
  onActionSuccess: (updatedCase: OnboardingCase) => void;
}

export const WorkflowPanel: React.FC<WorkflowPanelProps> = ({
  activeCase,
  currentUser,
  onActionSuccess
}) => {
  const [action, setAction] = useState<'Approve' | 'Return' | 'Reject' | null>(null);
  const [comments, setComments] = useState('');
  const [reviewerSign, setReviewerSign] = useState('');
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [error, setError] = useState('');
  const [previewPdf, setPreviewPdf] = useState<{ url: string; name: string } | null>(null);

  const returnActionLabel = useMemo(() => {
    if (currentUser.role === 'Compliance') return 'Return to Preparer';
    if (currentUser.role === 'EngagementPartner') return 'Return to Compliance';
    if (currentUser.role === 'RiskPartner') return 'Return to Engagement Partner';
    return 'Return';
  }, [currentUser.role]);

  const groupedEvidence = useMemo(() => {
    const docs = activeCase.documents || [];
    const inCategory = (...cats: string[]) => docs.filter((d) => cats.includes(d.category));
    return {
      idsAndDirectors: inCategory('ID_Passport', 'CR6_Directors', 'Incorporation_Doc', 'Trust_Deed', 'BO_Declaration'),
      cdd: inCategory('ID_Passport', 'Incorporation_Doc', 'CR6_Directors', 'Trust_Deed', 'BO_Declaration'),
      edd: inCategory('EDD_Funds_Proof', 'PEP_Mitigation'),
      all: docs,
    };
  }, [activeCase.documents]);

  // Determine if the current user is the authorized handler for this case
  const isAuthorizedHandler = useMemo(() => {
    // If case is not pending reviews, no action allowed
    if (activeCase.status === 'Draft' || activeCase.status === 'Closed' || activeCase.status === 'Rejected') {
      return false;
    }

    // Role verification
    if (
      (activeCase.status === 'Pending Compliance' || activeCase.status === 'ReturnedToCompliance') &&
      currentUser.role === 'Compliance'
    ) {
      return true;
    }
    if (
      (activeCase.status === 'Pending Risk Partner') &&
      currentUser.role === 'RiskPartner'
    ) {
      return true;
    }
    if (
      (activeCase.status === 'Pending Engagement Partner' || activeCase.status === 'ReturnedToEP') &&
      currentUser.role === 'EngagementPartner'
    ) {
      return true;
    }

    return false;
  }, [activeCase.status, activeCase.currentHandler, currentUser.role, currentUser.email]);

  // Submit Review Action
  const handleSubmitDecision = async () => {
    setError('');
    
    if (!action) {
      setError("Please select an action (Approve, Return, or Reject).");
      return;
    }

    // Mandatory comments check for Return or Reject
    if ((action === 'Return' || action === 'Reject') && !comments.trim()) {
      setError(`Mandatory comments are required to execute a client ${action.toLowerCase()} action.`);
      return;
    }

    if (!reviewerSign.trim()) {
      setError("Electronic signature and sign-off name is required.");
      return;
    }

    setIsSubmitting(true);
    try {
      const updated = await sharePointService.processWorkflowAction(
        activeCase.id,
        action,
        comments,
        currentUser.name,
        currentUser.role,
        reviewerSign
      );
      
      setIsSubmitting(false);
      setAction(null);
      setComments('');
      setReviewerSign('');
      alert(`Workflow transaction registered: Client portfolio ${action.toLowerCase()}d!`);
      onActionSuccess(updated);
    } catch (e) {
      setIsSubmitting(false);
      setError('Workflow error: ' + e);
    }
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: '1.5rem' }}>
      {previewPdf && (
        <div
          style={{
            position: 'fixed',
            inset: 0,
            backgroundColor: 'rgba(0,0,0,0.55)',
            zIndex: 1200,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            padding: '1rem',
          }}
          onClick={() => setPreviewPdf(null)}
        >
          <div
            className="card"
            style={{
              width: 'min(1100px, 96vw)',
              height: 'min(760px, 92vh)',
              padding: '0.75rem',
              display: 'flex',
              flexDirection: 'column',
              gap: '0.5rem',
            }}
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex-between" style={{ alignItems: 'center' }}>
              <div style={{ fontSize: '0.85rem', fontWeight: 700 }}>{previewPdf.name}</div>
              <div style={{ display: 'flex', gap: '0.5rem' }}>
                <button
                  className="btn btn-secondary"
                  onClick={() => window.open(previewPdf.url, '_blank', 'noopener,noreferrer')}
                >
                  Open in new tab
                </button>
                <button className="btn btn-secondary" onClick={() => setPreviewPdf(null)}>
                  Close
                </button>
              </div>
            </div>
            <iframe
              title={`PDF preview ${previewPdf.name}`}
              src={previewPdf.url}
              style={{
                width: '100%',
                height: '100%',
                border: '1px solid var(--color-border)',
                borderRadius: '8px',
                backgroundColor: 'white',
              }}
            />
            <span style={{ fontSize: '0.72rem', color: 'var(--color-text-muted)' }}>
              If this SharePoint PDF cannot render in-frame due to browser security policy, use "Open in new tab".
            </span>
          </div>
        </div>
      )}
      
      {/* Active review header */}
      {isAuthorizedHandler ? (
        <div className="card" style={{ borderLeft: '4px solid var(--color-primary)', backgroundColor: 'var(--color-primary-light)', padding: '1.25rem' }}>
          <h3 style={{ fontSize: '1rem', fontWeight: 700, color: 'var(--color-primary)', display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
            <UserCheck size={20} />
            Onboarding Review Underway
          </h3>
          <p style={{ fontSize: '0.8rem', color: 'var(--color-text-secondary)', marginTop: '0.25rem', lineHeight: '1.4' }}>
            As the designated <strong>{currentUser.role} Officer</strong>, you are authorized to verify the compliance materials, audit logs, and risk matrix. Execute sign-off or return to the engagement preparer.
          </p>

          {/* review selection buttons */}
          <div style={{ display: 'flex', gap: '0.75rem', marginTop: '1rem' }}>
            <button 
              className="btn btn-success" 
              style={{ padding: '0.5rem 1rem', fontSize: '0.8rem', opacity: action === 'Approve' ? 1 : 0.7 }}
              onClick={() => { setAction('Approve'); setError(''); }}
            >
              <CheckCircle size={14} /> Approve Portfolio
            </button>

            <button 
              className="btn btn-secondary" 
              style={{ 
                padding: '0.5rem 1rem', 
                fontSize: '0.8rem', 
                backgroundColor: action === 'Return' ? 'var(--color-returned-bg)' : 'transparent',
                borderColor: action === 'Return' ? 'var(--color-returned)' : 'var(--color-border)',
                color: action === 'Return' ? 'var(--color-returned)' : 'var(--color-text-secondary)',
                opacity: action === 'Return' ? 1 : 0.7 
              }}
              onClick={() => { setAction('Return'); setError(''); }}
            >
              <ArrowLeftRight size={14} /> {returnActionLabel}
            </button>

            <button 
              className="btn btn-danger" 
              style={{ padding: '0.5rem 1rem', fontSize: '0.8rem', opacity: action === 'Reject' ? 1 : 0.7 }}
              onClick={() => { setAction('Reject'); setError(''); }}
            >
              <XCircle size={14} /> Reject Client
            </button>
          </div>

          {/* Conditional inputs based on selection */}
          {action && (
            <div style={{ marginTop: '1.25rem', padding: '1rem', backgroundColor: 'var(--bg-card)', borderRadius: '8px', border: '1px solid var(--color-border)', display: 'flex', flexDirection: 'column', gap: '1rem' }}>
              
              <div className="form-group" style={{ margin: 0 }}>
                <label className="form-label">
                  Reviewer Decision Remarks
                  {(action === 'Return' || action === 'Reject') && <span className="required"> * (Mandatory for returns/rejections)</span>}
                </label>
                <textarea
                  rows={3}
                  className="form-control"
                  style={{ fontSize: '0.825rem' }}
                  placeholder={action === 'Approve' ? "Add structural comments or compliance approval notes..." : "Specify exact missing items or compliance flags requiring correction..."}
                  value={comments}
                  onChange={e => setComments(e.target.value)}
                />
              </div>

              <div className="form-group" style={{ margin: 0 }}>
                <label className="form-label">Electronic Signature / Sign-off Name <span className="required">*</span></label>
                <input
                  type="text"
                  className="form-control"
                  style={{ fontSize: '0.825rem', fontStyle: 'italic', fontWeight: 600 }}
                  placeholder="Enter full legal name to stamp signature..."
                  value={reviewerSign}
                  onChange={e => setReviewerSign(e.target.value)}
                />
              </div>

              {error && <span style={{ color: 'var(--color-high-risk)', fontSize: '0.75rem', fontWeight: 600 }}>{error}</span>}

              <button 
                className="btn btn-primary" 
                style={{ alignSelf: 'flex-end' }} 
                onClick={handleSubmitDecision}
                disabled={isSubmitting}
              >
                {isSubmitting ? 'Syncing Workflow...' : 'Submit Verification'}
              </button>
            </div>
          )}

        </div>
      ) : (
        <div className="card" style={{ padding: '1rem', display: 'flex', alignItems: 'center', gap: '0.5rem', backgroundColor: 'var(--bg-app)' }}>
          <Clock size={16} style={{ color: 'var(--color-text-muted)' }} />
          <span style={{ fontSize: '0.75rem', color: 'var(--color-text-secondary)', fontWeight: 550 }}>
            {activeCase.status === 'Closed' ? (
              <span style={{ color: 'var(--color-low-risk)', display: 'flex', alignItems: 'center', gap: '0.25rem' }}>
                <ShieldCheck size={14} /> Onboarding Portfolio is APPROVED & READ-ONLY (Locked under SharePoint Retention labels)
              </span>
            ) : activeCase.status === 'Rejected' ? (
              <span style={{ color: 'var(--color-high-risk)' }}>Onboarding portfolio was REJECTED by Compliance.</span>
            ) : (
              <>
                Current handler: <strong>{activeCase.currentHandler || '—'}</strong>.
                {' '}Assigned reviewers — Compliance: {activeCase.workflowAssignees.complianceEmail};
                {' '}Engagement Partner: {activeCase.workflowAssignees.engagementPartnerEmail};
                {' '}Risk Partner: {activeCase.workflowAssignees.riskPartnerEmail}.
              </>
            )}
          </span>
        </div>
      )}

      {/* Audit Logs Timeline Panel */}
      <div className="card">
        <h3 style={{ fontSize: '0.9rem', fontWeight: 700, textTransform: 'uppercase', color: 'var(--color-text-secondary)', letterSpacing: '0.05em', marginBottom: '1rem' }}>
          Reviewer Evidence View
        </h3>

        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: '0.75rem' }}>
          {[
            {
              title: 'Directors, IDs & BO Evidence',
              subtitle: `Directors captured: ${activeCase.directors.length}`,
              files: groupedEvidence.idsAndDirectors,
            },
            {
              title: 'CDD Evidence',
              subtitle: `CDD checks selected: ${Object.values(activeCase.cddMeasures).filter(Boolean).length}/6`,
              files: groupedEvidence.cdd,
            },
            {
              title: 'EDD / Source of Funds Evidence',
              subtitle: `EDD checks selected: ${Object.values(activeCase.eddApplied).filter(Boolean).length}/6`,
              files: groupedEvidence.edd,
            },
          ].map((section) => (
            <div key={section.title} style={{ border: '1px solid var(--color-border)', borderRadius: '8px', padding: '0.75rem', backgroundColor: 'var(--bg-app)' }}>
              <div style={{ fontWeight: 700, fontSize: '0.8rem', marginBottom: '0.2rem' }}>{section.title}</div>
              <div style={{ fontSize: '0.72rem', color: 'var(--color-text-muted)', marginBottom: '0.5rem' }}>{section.subtitle}</div>
              {section.files.length === 0 ? (
                <div style={{ fontSize: '0.75rem', color: 'var(--color-text-muted)' }}>No files in this section.</div>
              ) : (
                section.files.map((doc) => (
                  <div key={`${section.title}-${doc.id}`} style={{ fontSize: '0.75rem', marginBottom: '0.35rem', lineHeight: 1.35 }}>
                    {doc.webUrl ? (
                      <button
                        className="btn"
                        style={{
                          padding: 0,
                          border: 0,
                          background: 'transparent',
                          color: 'var(--color-primary)',
                          fontWeight: 600,
                          textDecoration: 'underline',
                          cursor: 'pointer',
                        }}
                        onClick={() => {
                          if (doc.name.toLowerCase().endsWith('.pdf')) {
                            setPreviewPdf({ url: doc.webUrl!, name: doc.name });
                          } else {
                            window.open(doc.webUrl, '_blank', 'noopener,noreferrer');
                          }
                        }}
                      >
                        {doc.name}
                      </button>
                    ) : (
                      <span style={{ fontWeight: 600 }}>{doc.name}</span>
                    )}
                    <div style={{ color: 'var(--color-text-muted)' }}>
                      {doc.category} • {doc.size} • {doc.uploadDate} • {doc.uploadedBy || 'Unknown uploader'}
                    </div>
                  </div>
                ))
              )}
            </div>
          ))}
        </div>

        <div style={{ marginTop: '0.75rem', fontSize: '0.75rem', color: 'var(--color-text-muted)' }}>
          Total attached documents: <strong>{groupedEvidence.all.length}</strong>
        </div>
      </div>

      <div className="card">
        <h3 style={{ fontSize: '0.9rem', fontWeight: 700, textTransform: 'uppercase', color: 'var(--color-text-secondary)', letterSpacing: '0.05em', marginBottom: '1.25rem', display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
          <MessageSquare size={16} style={{ color: 'var(--color-primary)' }} />
          Compliance Audit Timeline Trails
        </h3>

        <div className="timeline">
          {activeCase.auditLogs.map((log) => {
            let dotClass = 'timeline-dot';
            if (log.action === 'APPROVED' || log.action === 'SUBMITTED') dotClass = 'timeline-dot success';
            if (log.action === 'RETURNED') dotClass = 'timeline-dot returned';
            if (log.action === 'REJECTED') dotClass = 'timeline-dot error'; // mapped to timeline styling

            return (
              <div key={log.id} className="timeline-item">
                <div className={`${dotClass} ${log.id === 'log-seed' ? 'active' : ''}`} />
                <div className="timeline-content">
                  <div className="flex-between">
                    <span className="timeline-title" style={{ color: 'var(--color-text-main)' }}>
                      {log.action} &bull; <span style={{ fontWeight: 400, color: 'var(--color-text-secondary)' }}>{log.role}</span>
                    </span>
                    <span className="timeline-time">
                      {new Date(log.timestamp).toLocaleString()}
                    </span>
                  </div>
                  <span className="timeline-desc" style={{ color: 'var(--color-primary)', fontWeight: 550, fontSize: '0.75rem' }}>
                    Actor: {log.actor}
                  </span>
                  {log.comments && (
                    <p style={{ 
                      fontSize: '0.775rem', 
                      backgroundColor: 'var(--bg-app)', 
                      padding: '0.5rem 0.75rem', 
                      borderRadius: '6px', 
                      marginTop: '0.25rem',
                      borderLeft: '2px solid var(--color-border)',
                      lineHeight: '1.4' 
                    }}>
                      "{log.comments}"
                    </p>
                  )}
                </div>
              </div>
            );
          })}
        </div>
      </div>

    </div>
  );
};
