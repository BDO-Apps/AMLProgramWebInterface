import React, { useState, useMemo } from 'react';
import { 
  FileText, ChevronRight, ChevronLeft, Save, 
  Send, Plus, Trash2, Upload, ShieldCheck, Check, Info
} from 'lucide-react';
import { getDefaultWorkflowAssignees, isValidEmail } from '../config/workflowDefaults';
import { sharePointService } from '../services/SharePointService';
import type { OnboardingCase, Director, BeneficialOwner, DocumentUpload } from '../services/SharePointService';
import type { M365User } from '../services/MSALService';

interface OnboardingFormProps {
  caseId: string | null; // Null means create new
  currentUser: M365User;
  onClose: () => void;
  onSaveSuccess: (updatedCase: OnboardingCase) => void;
}

export const OnboardingForm: React.FC<OnboardingFormProps> = ({
  caseId,
  currentUser,
  onClose,
  onSaveSuccess
}) => {
  const [activeStep, setActiveStep] = useState(1);
  const [isUploading, setIsUploading] = useState(false);
  const [uploadProgress, setUploadProgress] = useState(0);
  const [errors, setErrors] = useState<Record<string, string>>({});
  
  // Signature Pad state
  const [preparerSignature, setPreparerSignature] = useState('');

  // Initial State Definition
  const [formData, setFormData] = useState<OnboardingCase>(() => {
    // If opening an existing case, we'll fetch in useEffect, but initialize placeholder first
    return {
      id: '',
      clientName: '',
      clientType: 'Legal Person',
      regNumber: '',
      registeredAddress: '',
      natureOfBusiness: '',
      purposeOfEngagement: '',
      contactInfo: { address: '', email: '', phone: '' },
      directors: [],
      beneficialOwners: [],
      hasBeneficialOwners: false,
      riskRatings: { client: 'Low', geography: 'Low', productService: 'Low', deliveryChannel: 'Low', paymentMode: 'Low' },
      overallRiskRating: 'Low',
      riskRationale: '',
      cddMeasures: { identityVerified: false, boVerified: false, natureUnderstood: false, pepScreened: false, sanctionsScreened: false, adverseMediaScreened: false },
      eddApplied: { sourceOfFundsVerified: false, sourceOfWealthVerified: false, enhancedAdverseMedia: false, additionalBoVerification: false, seniorMgmtApproved: false, enhancedMonitoringApplied: false },
      eddFindings: '',
      pepStatus: { isPep: false },
      sanctionsCheck: { screened: false, hasMatch: false },
      adverseMediaCheck: { hasAdverseInfo: false },
      decision: '',
      reviewFrequency: '',
      nextReviewDate: '',
      signatures: {},
      status: 'Draft',
      currentHandler: currentUser.email,
      workflowAssignees: getDefaultWorkflowAssignees(),
      documents: [],
      auditLogs: [],
      office: currentUser.office,
      dateCreated: '',
      lastUpdated: ''
    };
  });

  // Fetch case details if in edit mode
  React.useEffect(() => {
    if (caseId) {
      sharePointService.getCaseById(caseId).then(fetched => {
        if (fetched) {
          setFormData({
            ...fetched,
            workflowAssignees: fetched.workflowAssignees ?? getDefaultWorkflowAssignees(),
          });
          if (fetched.signatures?.preparer) {
            setPreparerSignature(fetched.signatures.preparer.sign);
          }
        }
      });
    }
  }, [caseId]);

  // Is Form read-only? (Read-only for everyone after submission, or read-only for non-preparers)
  const isReadOnly = useMemo(() => {
    if (formData.status !== 'Draft' && formData.status !== 'Returned') {
      return true;
    }
    return currentUser.role !== 'Preparer';
  }, [formData.status, currentUser.role]);

  // Handle standard field input
  const handleChange = (e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>) => {
    const { name, value } = e.target;
    
    // Support nested contact details
    if (name.startsWith('contact.')) {
      const field = name.split('.')[1];
      setFormData(prev => ({
        ...prev,
        contactInfo: { ...prev.contactInfo, [field]: value }
      }));
    } else {
      setFormData(prev => ({ ...prev, [name]: value }));
    }

    // Clear matching validation error
    if (errors[name]) {
      setErrors(prev => {
        const copy = { ...prev };
        delete copy[name];
        return copy;
      });
    }
  };

  // --- Dynamic EDD Trigger Calculation ---
  const isEddRequired = useMemo(() => {
    return (
      formData.overallRiskRating === 'High' || 
      formData.overallRiskRating === 'Medium' || 
      formData.pepStatus.isPep
    );
  }, [formData.overallRiskRating, formData.pepStatus.isPep]);

  // --- Relational Table Actions (Directors) ---
  const [newDirector, setNewDirector] = useState<Director>({ fullName: '', position: '', nationality: '', idNumber: '', countryOfResidence: '' });
  
  const addDirector = () => {
    if (!newDirector.fullName || !newDirector.position) {
      alert("Name and Position are required to register a director.");
      return;
    }
    setFormData(prev => ({
      ...prev,
      directors: [...prev.directors, { ...newDirector, id: 'dir-' + Date.now() }]
    }));
    setNewDirector({ fullName: '', position: '', nationality: '', idNumber: '', countryOfResidence: '' });
  };

  const removeDirector = (id: string) => {
    setFormData(prev => ({
      ...prev,
      directors: prev.directors.filter(d => d.id !== id)
    }));
  };

  // --- Relational Table Actions (Beneficial Owners) ---
  const [newBO, setNewBO] = useState<BeneficialOwner>({ fullName: '', ownershipPercentage: 25, basisOfControl: 'Direct Shares', country: '', verificationSource: 'Company Registry' });
  
  const addBO = () => {
    if (!newBO.fullName || newBO.ownershipPercentage <= 0) {
      alert("Valid Name and ownership percentage are required.");
      return;
    }
    setFormData(prev => ({
      ...prev,
      beneficialOwners: [...prev.beneficialOwners, { ...newBO, id: 'bo-' + Date.now() }],
      hasBeneficialOwners: true
    }));
    setNewBO({ fullName: '', ownershipPercentage: 25, basisOfControl: 'Direct Shares', country: '', verificationSource: 'Company registry' });
  };

  const removeBO = (id: string) => {
    const updatedBOs = formData.beneficialOwners.filter(b => b.id !== id);
    setFormData(prev => ({
      ...prev,
      beneficialOwners: updatedBOs,
      hasBeneficialOwners: updatedBOs.length > 0
    }));
  };

  // --- Interactive 3x3 Risk Matrix Matrix Calculation ---
  // Returns suggested risk score based on coordinate
  const suggestedRiskRating = useMemo(() => {
    const ratings = [
      formData.riskRatings.client,
      formData.riskRatings.geography,
      formData.riskRatings.productService,
      formData.riskRatings.deliveryChannel,
      formData.riskRatings.paymentMode
    ];
    
    const countHigh = ratings.filter(r => r === 'High').length;
    const countMed = ratings.filter(r => r === 'Medium').length;

    // Standard compliance aggregation:
    // If 2 or more metrics are HIGH, suggest HIGH.
    // If none are HIGH, but 2 or more are MEDIUM, suggest MEDIUM.
    // Otherwise LOW.
    if (countHigh >= 2) return 'High';
    if (countHigh === 1 || countMed >= 2) return 'Medium';
    return 'Low';
  }, [formData.riskRatings]);

  // Adjust specific inherent risk rating categories
  const setRiskCategory = (category: keyof OnboardingCase['riskRatings'], value: 'Low' | 'Medium' | 'High') => {
    if (isReadOnly) return;
    setFormData(prev => {
      const updatedRatings = { ...prev.riskRatings, [category]: value };
      
      // Auto-suggest overall risk based on category weightings
      const ratingsArr = Object.values(updatedRatings);
      const countHigh = ratingsArr.filter(r => r === 'High').length;
      const countMed = ratingsArr.filter(r => r === 'Medium').length;
      let overall: 'Low' | 'Medium' | 'High' = 'Low';
      if (countHigh >= 2) overall = 'High';
      else if (countHigh === 1 || countMed >= 2) overall = 'Medium';

      return {
        ...prev,
        riskRatings: updatedRatings,
        overallRiskRating: overall
      };
    });
  };

  // --- Document Upload Simulation ---
  const handleDocUpload = async (e: React.ChangeEvent<HTMLInputElement>, category: DocumentUpload['category']) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsUploading(true);
    setUploadProgress(10);
    
    // Simulate upload intervals
    const interval = setInterval(() => {
      setUploadProgress(prev => {
        if (prev >= 90) {
          clearInterval(interval);
          return 90;
        }
        return prev + 25;
      });
    }, 200);

    try {
      const uploaded = await sharePointService.uploadDocument(formData.id, file, category, currentUser.name);
      clearInterval(interval);
      setUploadProgress(100);
      setTimeout(() => {
        setFormData(prev => ({
          ...prev,
          documents: [...prev.documents, uploaded]
        }));
        setIsUploading(false);
        setUploadProgress(0);
      }, 300);
    } catch (err) {
      clearInterval(interval);
      setIsUploading(false);
      alert('Upload failed: ' + err);
    }
  };

  // --- Form Validation Safeguards ---
  const validateForm = (isFinalSubmit: boolean = false): boolean => {
    const newErrors: Record<string, string> = {};

    if (!formData.clientName.trim()) newErrors.clientName = 'Full Legal Name is required';
    if (!formData.regNumber.trim()) newErrors.regNumber = 'Registration or ID Number is required';
    if (!formData.natureOfBusiness.trim()) newErrors.natureOfBusiness = 'Nature of business is required';
    
    // Validation on Step 3 (Risk)
    if (isFinalSubmit && !formData.riskRationale.trim()) {
      newErrors.riskRationale = 'Mandatory risk rationale must explain inherent risk drivers';
    }

    // Document validation checks on submit
    if (isFinalSubmit) {
      if (formData.clientType === 'Individual' && !formData.documents.some(d => d.category === 'ID_Passport')) {
        newErrors.documents = 'Individual onboarding requires a verified ID/Passport document upload.';
      }
      if (formData.clientType === 'Legal Person' && !formData.documents.some(d => d.category === 'Incorporation_Doc')) {
        newErrors.documents = 'Corporate legal entity requires a Certificate of Incorporation upload.';
      }
      if (formData.clientType === 'Legal Person' && !formData.documents.some(d => d.category === 'CR6_Directors')) {
        newErrors.documents = 'Zimbabwe incorporation guidelines mandate CR6 Directors Registry files.';
      }
      if (formData.clientType === 'Legal Arrangement' && !formData.documents.some(d => d.category === 'Trust_Deed')) {
        newErrors.documents = 'Trust structures require a validated Trust Deed upload.';
      }
      if (isEddRequired && !formData.documents.some(d => d.category === 'EDD_Funds_Proof')) {
        newErrors.documents = 'EDD triggered cases require Source of Wealth or Source of Funds documentation.';
      }
      if (!preparerSignature.trim()) {
        newErrors.signature = 'Engagement preparer signature and sign-off name is required.';
      }
      const { complianceEmail, engagementPartnerEmail, riskPartnerEmail } = formData.workflowAssignees;
      if (!isValidEmail(complianceEmail)) {
        newErrors.complianceReviewerEmail = 'A valid Compliance / MLRO reviewer email is required.';
      }
      if (!isValidEmail(engagementPartnerEmail)) {
        newErrors.engagementPartnerReviewerEmail = 'A valid Engagement Partner reviewer email is required.';
      }
      if (!isValidEmail(riskPartnerEmail)) {
        newErrors.riskPartnerReviewerEmail = 'A valid Risk Partner reviewer email is required.';
      }
    }

    setErrors(newErrors);
    
    // Highlight first step with errors
    if (Object.keys(newErrors).length > 0) {
      if (newErrors.clientName || newErrors.regNumber || newErrors.natureOfBusiness) {
        setActiveStep(1);
      } else if (newErrors.riskRationale) {
        setActiveStep(3);
      } else if (newErrors.documents) {
        setActiveStep(4);
      } else if (
        newErrors.signature ||
        newErrors.complianceReviewerEmail ||
        newErrors.engagementPartnerReviewerEmail ||
        newErrors.riskPartnerReviewerEmail
      ) {
        setActiveStep(7);
      }
      return false;
    }
    return true;
  };

  // --- Save / Draft submission ---
  const handleSaveDraft = async () => {
    if (!validateForm(false)) return;
    try {
      const saved = await sharePointService.saveCase(
        formData,
        currentUser.email,
        currentUser.role
      );
      setFormData(saved);
      onSaveSuccess(saved);
      alert(
        saved.directors.length || saved.beneficialOwners.length
          ? `Draft saved (${saved.id}) with ${saved.directors.length} director(s) and ${saved.beneficialOwners.length} beneficial owner(s).`
          : `Draft saved (${saved.id}). Save again after adding ownership records on Step 2.`
      );
    } catch (e) {
      alert('Error saving draft: ' + e);
    }
  };

  // --- Final Submit Action ---
  const handleFinalSubmit = async () => {
    if (!validateForm(true)) return;
    if (window.confirm("Submit this client onboarding portfolio to compliance for review? This locks editing controls.")) {
      try {
        const assignees = {
          complianceEmail: formData.workflowAssignees.complianceEmail.trim().toLowerCase(),
          engagementPartnerEmail: formData.workflowAssignees.engagementPartnerEmail.trim().toLowerCase(),
          riskPartnerEmail: formData.workflowAssignees.riskPartnerEmail.trim().toLowerCase(),
        };
        const submitted = await sharePointService.submitCase(
          formData.id,
          currentUser.name,
          currentUser.role,
          preparerSignature,
          assignees
        );
        setFormData(submitted);
        onSaveSuccess(submitted);
        alert('Case portfolio successfully submitted to BDO Compliance. SharePoint workflow notifications activated.');
        onClose();
      } catch (e) {
        alert('Error during submission: ' + e);
      }
    }
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: '1.5rem' }}>
      
      {/* Back Header */}
      <div className="flex-between">
        <div>
          <h2 style={{ fontSize: '1.5rem', fontWeight: 700, color: 'var(--color-text-main)', display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
            {formData.id ? `Case: ${formData.id}` : 'Create Client Case'}
            <span className="tagline" style={{ textTransform: 'none' }}>{formData.status}</span>
          </h2>
          <p style={{ fontSize: '0.85rem', color: 'var(--color-text-secondary)' }}>
            Client Onboarding RBA Portfolio &bull; BDO {formData.office}
          </p>
        </div>
        
        <button className="btn btn-secondary" onClick={onClose}>
          Exit Portal
        </button>
      </div>

      {/* stepper */}
      <div className="stepper">
        {[
          { step: 1, label: 'Identification' },
          { step: 2, label: 'Ownership' },
          { step: 3, label: 'Inherent Risk' },
          { step: 4, label: 'CDD Vault' },
          { step: 5, label: 'EDD Actions', disabled: !isEddRequired },
          { step: 6, label: 'Screening' },
          { step: 7, label: 'Sign-off' }
        ].map(item => (
          <div 
            key={item.step} 
            className={`step-item ${activeStep === item.step ? 'active' : ''} ${activeStep > item.step ? 'completed' : ''}`}
            onClick={() => {
              if (!item.disabled) setActiveStep(item.step);
            }}
            style={{ opacity: item.disabled ? 0.35 : 1, cursor: item.disabled ? 'not-allowed' : 'pointer' }}
          >
            <div className="step-bubble">
              {activeStep > item.step ? <Check size={16} /> : item.step}
            </div>
            <div className="step-label">{item.label}</div>
          </div>
        ))}
      </div>

      {/* Main Form Fields Panel */}
      <div className="card" style={{ padding: '2rem' }}>
        
        {/* Step 1: Identification */}
        {activeStep === 1 && (
          <div style={{ display: 'flex', flexDirection: 'column', gap: '1.25rem' }}>
            <h3 style={{ fontSize: '1.1rem', fontWeight: 700, borderBottom: '2px solid var(--color-border)', paddingBottom: '0.5rem', marginBottom: '0.5rem' }}>
              Step 1: Client Identification (Basic CDD Information)
            </h3>
            
            <div className="form-group">
              <label className="form-label">Client Category Type <span className="required">*</span></label>
              <div style={{ display: 'flex', gap: '1.5rem', marginTop: '0.25rem' }}>
                {['Individual', 'Legal Person', 'Legal Arrangement'].map(type => (
                  <label key={type} style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', cursor: isReadOnly ? 'not-allowed' : 'pointer' }}>
                    <input 
                      type="radio" 
                      name="clientType" 
                      value={type} 
                      checked={formData.clientType === type}
                      onChange={handleChange}
                      disabled={isReadOnly}
                    />
                    <span style={{ fontSize: '0.9rem', fontWeight: 500 }}>{type}</span>
                  </label>
                ))}
              </div>
            </div>

            <div className="grid-2">
              <div className="form-group">
                <label className="form-label">Legal / Full Name <span className="required">*</span></label>
                <input 
                  type="text" 
                  name="clientName" 
                  className={`form-control ${errors.clientName ? 'error' : ''}`}
                  value={formData.clientName} 
                  onChange={handleChange} 
                  disabled={isReadOnly}
                  placeholder="e.g. Kariba Minerals Ltd or John Doe"
                />
                {errors.clientName && <span style={{ color: 'var(--color-high-risk)', fontSize: '0.75rem' }}>{errors.clientName}</span>}
              </div>

              <div className="form-group">
                <label className="form-label">
                  {formData.clientType === 'Individual' ? 'ID / Passport Number' : 'Company Registration / Trust Number'} <span className="required">*</span>
                </label>
                <input 
                  type="text" 
                  name="regNumber" 
                  className="form-control" 
                  value={formData.regNumber} 
                  onChange={handleChange} 
                  disabled={isReadOnly}
                  placeholder="e.g. CO-8821 or 63-9992345"
                />
                {errors.regNumber && <span style={{ color: 'var(--color-high-risk)', fontSize: '0.75rem' }}>{errors.regNumber}</span>}
              </div>
            </div>

            <div className="grid-2">
              <div className="form-group">
                <label className="form-label">Nature of Business / Industry Sector <span className="required">*</span></label>
                <input 
                  type="text" 
                  name="natureOfBusiness" 
                  className="form-control" 
                  value={formData.natureOfBusiness} 
                  onChange={handleChange} 
                  disabled={isReadOnly}
                  placeholder="e.g. Extractive Gem Mining or Retail Banking"
                />
                {errors.natureOfBusiness && <span style={{ color: 'var(--color-high-risk)', fontSize: '0.75rem' }}>{errors.natureOfBusiness}</span>}
              </div>

              <div className="form-group">
                <label className="form-label">Purpose of Engagement with BDO <span className="required">*</span></label>
                <input 
                  type="text" 
                  name="purposeOfEngagement" 
                  className="form-control" 
                  value={formData.purposeOfEngagement} 
                  onChange={handleChange} 
                  disabled={isReadOnly}
                  placeholder="e.g. Annual Audit & Tax Compliance Review"
                />
              </div>
            </div>

            <div className="form-group">
              <label className="form-label">Registered Principal Address <span className="required">*</span></label>
              <textarea 
                name="registeredAddress" 
                rows={2} 
                className="form-control" 
                value={formData.registeredAddress} 
                onChange={handleChange} 
                disabled={isReadOnly}
                placeholder="Physical street address..."
              />
            </div>

            <h4 style={{ fontSize: '0.95rem', fontWeight: 600, marginTop: '0.5rem', color: 'var(--color-primary)' }}>Contact Information</h4>
            <div className="grid-3">
              <div className="form-group">
                <label className="form-label">Contact Person Address</label>
                <input 
                  type="text" 
                  name="contact.address" 
                  className="form-control" 
                  value={formData.contactInfo.address} 
                  onChange={handleChange} 
                  disabled={isReadOnly}
                />
              </div>
              <div className="form-group">
                <label className="form-label">Official Email</label>
                <input 
                  type="email" 
                  name="contact.email" 
                  className="form-control" 
                  value={formData.contactInfo.email} 
                  onChange={handleChange} 
                  disabled={isReadOnly}
                />
              </div>
              <div className="form-group">
                <label className="form-label">Phone Line</label>
                <input 
                  type="text" 
                  name="contact.phone" 
                  className="form-control" 
                  value={formData.contactInfo.phone} 
                  onChange={handleChange} 
                  disabled={isReadOnly}
                />
              </div>
            </div>
          </div>
        )}

        {/* Step 2: Ownership */}
        {activeStep === 2 && (
          <div style={{ display: 'flex', flexDirection: 'column', gap: '1.5rem' }}>
            <h3 style={{ fontSize: '1.1rem', fontWeight: 700, borderBottom: '2px solid var(--color-border)', paddingBottom: '0.5rem' }}>
              Step 2: Ownership, Control & Management (FATF R.24 & R.25)
            </h3>

            {/* Directors Sheet */}
            <div>
              <h4 style={{ fontSize: '0.95rem', fontWeight: 700, color: 'var(--color-text-secondary)', marginBottom: '0.75rem', display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
                Directors / Trustees / Senior Management Panel
                <span style={{ fontSize: '0.75rem', fontWeight: 400, color: 'var(--color-text-muted)' }}>({formData.directors.length} registered)</span>
              </h4>
              
              <div className="table-wrapper" style={{ marginBottom: '1rem' }}>
                <table className="bdo-table" style={{ fontSize: '0.8rem' }}>
                  <thead>
                    <tr>
                      <th>Full Legal Name</th>
                      <th>Corporate Position</th>
                      <th>Nationality</th>
                      <th>ID / Passport Number</th>
                      <th>Residency</th>
                      {!isReadOnly && <th style={{ width: '50px' }}>Action</th>}
                    </tr>
                  </thead>
                  <tbody>
                    {formData.directors.length > 0 ? (
                      formData.directors.map(dir => (
                        <tr key={dir.id}>
                          <td style={{ fontWeight: 600 }}>{dir.fullName}</td>
                          <td>{dir.position}</td>
                          <td>{dir.nationality}</td>
                          <td>{dir.idNumber}</td>
                          <td>{dir.countryOfResidence}</td>
                          {!isReadOnly && (
                            <td>
                              <button className="btn" style={{ padding: '4px', color: 'var(--color-high-risk)' }} onClick={() => removeDirector(dir.id!)}>
                                <Trash2 size={14} />
                              </button>
                            </td>
                          )}
                        </tr>
                      ))
                    ) : (
                      <tr>
                        <td colSpan={6} style={{ textAlign: 'center', color: 'var(--color-text-muted)' }}>No board directors logged. Please register below.</td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>

              {!isReadOnly && (
                <div style={{ padding: '1rem', backgroundColor: 'var(--bg-app)', borderRadius: '8px', border: '1px solid var(--color-border)', display: 'flex', flexDirection: 'column', gap: '0.75rem' }}>
                  <span style={{ fontSize: '0.8rem', fontWeight: 700, color: 'var(--color-primary)' }}>+ Register Director / Management Member</span>
                  <div className="grid-4" style={{ gap: '0.5rem' }}>
                    <input type="text" placeholder="Name" className="form-control" style={{ fontSize: '0.8rem' }} value={newDirector.fullName} onChange={e => setNewDirector(p => ({ ...p, fullName: e.target.value }))} />
                    <input type="text" placeholder="Position (e.g. Managing Dir)" className="form-control" style={{ fontSize: '0.8rem' }} value={newDirector.position} onChange={e => setNewDirector(p => ({ ...p, position: e.target.value }))} />
                    <input type="text" placeholder="Nationality" className="form-control" style={{ fontSize: '0.8rem' }} value={newDirector.nationality} onChange={e => setNewDirector(p => ({ ...p, nationality: e.target.value }))} />
                    <div style={{ display: 'flex', gap: '0.25rem' }}>
                      <input type="text" placeholder="ID/Passport" className="form-control" style={{ fontSize: '0.8rem' }} value={newDirector.idNumber} onChange={e => setNewDirector(p => ({ ...p, idNumber: e.target.value }))} />
                      <input type="text" placeholder="Residency" className="form-control" style={{ fontSize: '0.8rem' }} value={newDirector.countryOfResidence} onChange={e => setNewDirector(p => ({ ...p, countryOfResidence: e.target.value }))} />
                      <button className="btn btn-primary" style={{ padding: '0.35rem 0.75rem' }} onClick={addDirector}>
                        <Plus size={16} />
                      </button>
                    </div>
                  </div>
                </div>
              )}
            </div>

            {/* Beneficial Owners Sheet */}
            <div style={{ marginTop: '1rem' }}>
              <div className="alert-panel" style={{ margin: '0 0 1rem', padding: '0.75rem', display: 'flex', gap: '0.5rem', alignItems: 'center' }}>
                <Info size={16} style={{ color: 'var(--color-primary)', flexShrink: 0 }} />
                <span style={{ fontSize: '0.75rem', lineHeight: '1.4' }}>
                  <strong>FATF R.24 Compliance Rule</strong>: Natural persons holding <strong>25% or more</strong> direct or indirect equity/voting interest must be identified and verified.
                </span>
              </div>

              <h4 style={{ fontSize: '0.95rem', fontWeight: 700, color: 'var(--color-text-secondary)', marginBottom: '0.75rem', display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
                Ultimate Beneficial Owners (UBO) Registry
                <span style={{ fontSize: '0.75rem', fontWeight: 400, color: 'var(--color-text-muted)' }}>({formData.beneficialOwners.length} registered)</span>
              </h4>

              <div className="table-wrapper" style={{ marginBottom: '1rem' }}>
                <table className="bdo-table" style={{ fontSize: '0.8rem' }}>
                  <thead>
                    <tr>
                      <th>UBO Legal Name</th>
                      <th>Equity / Control %</th>
                      <th>Basis of Control</th>
                      <th>Country</th>
                      <th>Verification Source</th>
                      {!isReadOnly && <th style={{ width: '50px' }}>Action</th>}
                    </tr>
                  </thead>
                  <tbody>
                    {formData.beneficialOwners.length > 0 ? (
                      formData.beneficialOwners.map(bo => (
                        <tr key={bo.id}>
                          <td style={{ fontWeight: 600 }}>{bo.fullName}</td>
                          <td style={{ fontWeight: 700, color: 'var(--color-primary)' }}>{bo.ownershipPercentage}%</td>
                          <td>{bo.basisOfControl}</td>
                          <td>{bo.country}</td>
                          <td>{bo.verificationSource}</td>
                          {!isReadOnly && (
                            <td>
                              <button className="btn" style={{ padding: '4px', color: 'var(--color-high-risk)' }} onClick={() => removeBO(bo.id!)}>
                                <Trash2 size={14} />
                              </button>
                            </td>
                          )}
                        </tr>
                      ))
                    ) : (
                      <tr>
                        <td colSpan={6} style={{ textAlign: 'center', color: 'var(--color-text-muted)' }}>No beneficial owners logged holding &ge; 25% ownership.</td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>

              {!isReadOnly && (
                <div style={{ padding: '1rem', backgroundColor: 'var(--bg-app)', borderRadius: '8px', border: '1px solid var(--color-border)', display: 'flex', flexDirection: 'column', gap: '0.75rem' }}>
                  <span style={{ fontSize: '0.8rem', fontWeight: 700, color: 'var(--color-primary)' }}>+ Add Ultimate Beneficial Owner (UBO)</span>
                  <div className="grid-3" style={{ gap: '0.5rem' }}>
                    <input type="text" placeholder="Name" className="form-control" style={{ fontSize: '0.8rem' }} value={newBO.fullName} onChange={e => setNewBO(p => ({ ...p, fullName: e.target.value }))} />
                    
                    <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
                      <span style={{ fontSize: '0.75rem', fontWeight: 600, width: '40px' }}>Hold %:</span>
                      <input type="number" placeholder="Percentage" min="1" max="100" className="form-control" style={{ fontSize: '0.8rem' }} value={newBO.ownershipPercentage} onChange={e => setNewBO(p => ({ ...p, ownershipPercentage: parseInt(e.target.value) || 25 }))} />
                    </div>

                    <input type="text" placeholder="Basis of Control (e.g. Voting Shares)" className="form-control" style={{ fontSize: '0.8rem' }} value={newBO.basisOfControl} onChange={e => setNewBO(p => ({ ...p, basisOfControl: e.target.value }))} />
                  </div>
                  
                  <div className="grid-3" style={{ gap: '0.5rem' }}>
                    <input type="text" placeholder="Country" className="form-control" style={{ fontSize: '0.8rem' }} value={newBO.country} onChange={e => setNewBO(p => ({ ...p, country: e.target.value }))} />
                    
                    <select className="form-control" style={{ fontSize: '0.8rem' }} value={newBO.verificationSource} onChange={e => setNewBO(p => ({ ...p, verificationSource: e.target.value }))}>
                      <option value="Company Registry">Company Registry / CIPA</option>
                      <option value="Trust Deed">Trust Deed</option>
                      <option value="Share Register">Share Register or certificates</option>
                      <option value="Beneficial ownership declaration">Beneficial ownership declaration</option>
                      <option value="Other">Other Verified Document</option>
                    </select>

                    <button className="btn btn-primary" style={{ height: '36px' }} onClick={addBO}>
                      <Plus size={16} /> Register Owner
                    </button>
                  </div>
                </div>
              )}
            </div>
          </div>
        )}

        {/* Step 3: Inherent Risk */}
        {activeStep === 3 && (
          <div style={{ display: 'flex', flexDirection: 'column', gap: '1.5rem' }}>
            <h3 style={{ fontSize: '1.1rem', fontWeight: 700, borderBottom: '2px solid var(--color-border)', paddingBottom: '0.5rem' }}>
              Step 3: Inherent Risk Assessment & Proportionality (FATF R.1)
            </h3>

            <div style={{ display: 'grid', gridTemplateColumns: '1.2fr 1fr', gap: '2rem' }}>
              
              {/* Category selections */}
              <div style={{ display: 'flex', flexDirection: 'column', gap: '1rem' }}>
                <span style={{ fontSize: '0.85rem', fontWeight: 700, color: 'var(--color-primary)' }}>Set Sector Risk Indicators</span>
                
                {[
                  { key: 'client', label: 'Client Risk Factors (PEP, complex structure, NPO)' },
                  { key: 'geography', label: 'Geographical Risk Factors (Greylist, corruption rate)' },
                  { key: 'productService', label: 'Product & Service Risk (Funds holding, formations)' },
                  { key: 'deliveryChannel', label: 'Delivery Channels (Non-face-to-face, intermediary)' },
                  { key: 'paymentMode', label: 'Payment Mode Risk (Cash transactions, virtual assets)' }
                ].map(item => (
                  <div key={item.key} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '0.5rem 0', borderBottom: '1px solid var(--color-border)' }}>
                    <span style={{ fontSize: '0.825rem', fontWeight: 600, color: 'var(--color-text-secondary)' }}>{item.label}</span>
                    <div style={{ display: 'flex', gap: '0.25rem' }}>
                      {['Low', 'Medium', 'High'].map(level => {
                        const active = formData.riskRatings[item.key as keyof OnboardingCase['riskRatings']] === level;
                        let activeColor = 'var(--color-low-risk)';
                        if (level === 'Medium') activeColor = 'var(--color-med-risk)';
                        if (level === 'High') activeColor = 'var(--color-high-risk)';

                        return (
                          <button
                            key={level}
                            className="btn"
                            disabled={isReadOnly}
                            style={{ 
                              padding: '0.25rem 0.5rem', 
                              fontSize: '0.75rem', 
                              borderRadius: '4px',
                              backgroundColor: active ? activeColor : 'var(--bg-app)',
                              color: active ? 'white' : 'var(--color-text-secondary)',
                              fontWeight: 600
                            }}
                            onClick={() => setRiskCategory(item.key as any, level as any)}
                          >
                            {level}
                          </button>
                        );
                      })}
                    </div>
                  </div>
                ))}
              </div>

              {/* suggested rating / visual matrix */}
              <div style={{ display: 'flex', flexDirection: 'column', gap: '1rem', alignItems: 'center', padding: '1rem', backgroundColor: 'var(--bg-app)', borderRadius: '12px', border: '1px solid var(--color-border)' }}>
                <h4 style={{ fontSize: '0.85rem', fontWeight: 700, color: 'var(--color-text-secondary)', alignSelf: 'flex-start' }}>Interactive Suggested Inherent Risk Matrix</h4>
                
                {/* 3x3 Likelihood (Y) vs Impact (X) Visual representation */}
                <div className="risk-matrix-wrapper">
                  <div className="risk-matrix">
                    <span className="matrix-label-y">Impact</span>
                    
                    {/* Rows from High down to Low */}
                    {/* Impact High row */}
                    <div className={`matrix-cell cell-med ${formData.overallRiskRating === 'Medium' ? 'active' : ''}`}>M</div>
                    <div className={`matrix-cell cell-high ${formData.overallRiskRating === 'High' ? 'active' : ''}`}>H</div>
                    <div className={`matrix-cell cell-high ${formData.overallRiskRating === 'High' ? 'active' : ''}`}>H</div>
                    
                    {/* Impact Medium row */}
                    <div className={`matrix-cell cell-low ${formData.overallRiskRating === 'Low' ? 'active' : ''}`}>L</div>
                    <div className={`matrix-cell cell-med ${formData.overallRiskRating === 'Medium' ? 'active' : ''}`}>M</div>
                    <div className={`matrix-cell cell-high ${formData.overallRiskRating === 'High' ? 'active' : ''}`}>H</div>
                    
                    {/* Impact Low row */}
                    <div className={`matrix-cell cell-low ${formData.overallRiskRating === 'Low' ? 'active' : ''}`}>L</div>
                    <div className={`matrix-cell cell-low ${formData.overallRiskRating === 'Low' ? 'active' : ''}`}>L</div>
                    <div className={`matrix-cell cell-med ${formData.overallRiskRating === 'Medium' ? 'active' : ''}`}>M</div>
                    
                    <span className="matrix-label-x">Likelihood</span>
                  </div>
                </div>

                <div style={{ textAlign: 'center', marginTop: '0.5rem' }}>
                  <span style={{ fontSize: '0.75rem', color: 'var(--color-text-muted)' }}>Aggregated Rationale suggests:</span>
                  <div style={{ fontSize: '1.25rem', fontWeight: 800, marginTop: '0.25rem', color: 'var(--color-primary)' }}>
                    {suggestedRiskRating} Risk suggested
                  </div>
                  
                  {formData.overallRiskRating !== suggestedRiskRating && (
                    <div style={{ fontSize: '0.7rem', color: 'var(--color-high-risk)', fontWeight: 600, marginTop: '0.25rem' }}>
                      * Professional Override Enabled ({formData.overallRiskRating} Risk selected)
                    </div>
                  )}
                </div>
              </div>

            </div>

            {/* Overall Inherent Risk Selection & Rationale */}
            <div style={{ borderTop: '1.5px solid var(--color-border)', paddingTop: '1.25rem', marginTop: '0.5rem' }}>
              <div className="grid-2" style={{ gridTemplateColumns: '200px 1fr' }}>
                <div className="form-group">
                  <label className="form-label">Overall Inherent Risk Rating <span className="required">*</span></label>
                  <select 
                    name="overallRiskRating" 
                    className="form-control" 
                    style={{ fontWeight: 700 }}
                    value={formData.overallRiskRating}
                    onChange={handleChange}
                    disabled={isReadOnly}
                  >
                    <option value="Low">Low Rating</option>
                    <option value="Medium">Medium Rating</option>
                    <option value="High">High Rating</option>
                  </select>
                </div>

                <div className="form-group">
                  <label className="form-label">Risk Rating Rationale (Mandatory Explanation) <span className="required">*</span></label>
                  <textarea 
                    name="riskRationale" 
                    rows={3} 
                    className="form-control" 
                    value={formData.riskRationale} 
                    onChange={handleChange} 
                    disabled={isReadOnly}
                    placeholder="Provide a solid compliance rationale explaining risk drivers (industry factors, geographic links, complexity override)..."
                  />
                  {errors.riskRationale && <span style={{ color: 'var(--color-high-risk)', fontSize: '0.75rem' }}>{errors.riskRationale}</span>}
                </div>
              </div>
            </div>

          </div>
        )}

        {/* Step 4: CDD Vault */}
        {activeStep === 4 && (
          <div style={{ display: 'flex', flexDirection: 'column', gap: '1.5rem' }}>
            <h3 style={{ fontSize: '1.1rem', fontWeight: 700, borderBottom: '2px solid var(--color-border)', paddingBottom: '0.5rem' }}>
              Step 4: Customer Due Diligence (CDD) Measures & Document Vault
            </h3>

            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1.2fr', gap: '2rem' }}>
              
              {/* Checklists */}
              <div style={{ display: 'flex', flexDirection: 'column', gap: '0.85rem' }}>
                <span style={{ fontSize: '0.85rem', fontWeight: 700, color: 'var(--color-primary)' }}>Standard CDD Requirements Checklist</span>
                
                {[
                  { key: 'identityVerified', label: 'Identity verified using reliable, independent source documents' },
                  { key: 'boVerified', label: 'Beneficial ownership identified and verified on a risk-based basis' },
                  { key: 'natureUnderstood', label: 'Purpose and intended nature of business relationship understood' },
                  { key: 'pepScreened', label: 'Politically Exposed Person (PEP) screening conducted' },
                  { key: 'sanctionsScreened', label: 'UNSCR Sanctions screening completed (100% clean record)' },
                  { key: 'adverseMediaScreened', label: 'Adverse and negative news media screening conducted' }
                ].map(item => (
                  <label key={item.key} style={{ display: 'flex', gap: '0.75rem', alignItems: 'center', cursor: isReadOnly ? 'not-allowed' : 'pointer', fontSize: '0.825rem' }}>
                    <input
                      type="checkbox"
                      checked={formData.cddMeasures[item.key as keyof OnboardingCase['cddMeasures']]}
                      disabled={isReadOnly}
                      onChange={(e) => {
                        const checked = e.target.checked;
                        setFormData(prev => ({
                          ...prev,
                          cddMeasures: { ...prev.cddMeasures, [item.key]: checked }
                        }));
                      }}
                    />
                    <span style={{ color: 'var(--color-text-secondary)', fontWeight: 500 }}>{item.label}</span>
                  </label>
                ))}
              </div>

              {/* File Vault */}
              <div>
                <span style={{ fontSize: '0.85rem', fontWeight: 700, color: 'var(--color-primary)', display: 'block', marginBottom: '0.75rem' }}>Compliance Document Upload Vault</span>

                {/* Upload checklist requirements */}
                <div style={{ display: 'flex', flexDirection: 'column', gap: '0.5rem', marginBottom: '1rem', padding: '0.75rem', backgroundColor: 'var(--bg-app)', borderRadius: '8px', border: '1px solid var(--color-border)' }}>
                  <span style={{ fontSize: '0.75rem', fontWeight: 700, color: 'var(--color-text-secondary)' }}>Required Files Status:</span>
                  
                  {formData.clientType === 'Individual' && (
                    <div style={{ display: 'flex', alignItems: 'center', gap: '0.35rem', fontSize: '0.725rem', fontWeight: 600 }}>
                      {formData.documents.some(d => d.category === 'ID_Passport') ? <Check size={12} color="var(--color-low-risk)" /> : <Info size={12} color="var(--color-high-risk)" />}
                      <span style={{ color: formData.documents.some(d => d.category === 'ID_Passport') ? 'var(--color-low-risk)' : 'var(--color-high-risk)' }}>ID / Passport (Individual/Directors)</span>
                    </div>
                  )}

                  {formData.clientType === 'Legal Person' && (
                    <>
                      <div style={{ display: 'flex', alignItems: 'center', gap: '0.35rem', fontSize: '0.725rem', fontWeight: 600 }}>
                        {formData.documents.some(d => d.category === 'Incorporation_Doc') ? <Check size={12} color="var(--color-low-risk)" /> : <Info size={12} color="var(--color-high-risk)" />}
                        <span style={{ color: formData.documents.some(d => d.category === 'Incorporation_Doc') ? 'var(--color-low-risk)' : 'var(--color-high-risk)' }}>Certificate of Incorporation / Articles</span>
                      </div>
                      <div style={{ display: 'flex', alignItems: 'center', gap: '0.35rem', fontSize: '0.725rem', fontWeight: 600 }}>
                        {formData.documents.some(d => d.category === 'CR6_Directors') ? <Check size={12} color="var(--color-low-risk)" /> : <Info size={12} color="var(--color-high-risk)" />}
                        <span style={{ color: formData.documents.some(d => d.category === 'CR6_Directors') ? 'var(--color-low-risk)' : 'var(--color-high-risk)' }}>CR6 Directors Registry Form</span>
                      </div>
                    </>
                  )}

                  {formData.clientType === 'Legal Arrangement' && (
                    <div style={{ display: 'flex', alignItems: 'center', gap: '0.35rem', fontSize: '0.725rem', fontWeight: 600 }}>
                      {formData.documents.some(d => d.category === 'Trust_Deed') ? <Check size={12} color="var(--color-low-risk)" /> : <Info size={12} color="var(--color-high-risk)" />}
                      <span style={{ color: formData.documents.some(d => d.category === 'Trust_Deed') ? 'var(--color-low-risk)' : 'var(--color-high-risk)' }}>Trust Deed Appointment Documents</span>
                    </div>
                  )}

                  {isEddRequired && (
                    <div style={{ display: 'flex', alignItems: 'center', gap: '0.35rem', fontSize: '0.725rem', fontWeight: 600 }}>
                      {formData.documents.some(d => d.category === 'EDD_Funds_Proof') ? <Check size={12} color="var(--color-low-risk)" /> : <Info size={12} color="var(--color-high-risk)" />}
                      <span style={{ color: formData.documents.some(d => d.category === 'EDD_Funds_Proof') ? 'var(--color-low-risk)' : 'var(--color-high-risk)' }}>EDD: Verified Source of Funds / Wealth Proof</span>
                    </div>
                  )}
                </div>

                {/* Upload drag drop simulator */}
                {!isReadOnly && (
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '0.75rem', marginBottom: '1rem' }}>
                    <div className="grid-2" style={{ gap: '0.5rem' }}>
                      <div className="file-uploader" style={{ padding: '1rem 0.5rem', position: 'relative' }}>
                        <Upload size={20} style={{ margin: '0 auto 0.25rem', display: 'block', color: 'var(--color-text-muted)' }} />
                        <span style={{ fontSize: '0.75rem', fontWeight: 600 }}>Upload ID / Incorporation</span>
                        <input 
                          type="file" 
                          style={{ position: 'absolute', top: 0, left: 0, right: 0, bottom: 0, opacity: 0, cursor: 'pointer' }}
                          onChange={(e) => handleDocUpload(e, formData.clientType === 'Individual' ? 'ID_Passport' : 'Incorporation_Doc')}
                        />
                      </div>

                      {formData.clientType === 'Legal Person' && (
                        <div className="file-uploader" style={{ padding: '1rem 0.5rem', position: 'relative' }}>
                          <Upload size={20} style={{ margin: '0 auto 0.25rem', display: 'block', color: 'var(--color-text-muted)' }} />
                          <span style={{ fontSize: '0.75rem', fontWeight: 600 }}>Upload Zimbabwe CR6</span>
                          <input 
                            type="file" 
                            style={{ position: 'absolute', top: 0, left: 0, right: 0, bottom: 0, opacity: 0, cursor: 'pointer' }}
                            onChange={(e) => handleDocUpload(e, 'CR6_Directors')}
                          />
                        </div>
                      )}

                      {formData.clientType === 'Legal Arrangement' && (
                        <div className="file-uploader" style={{ padding: '1rem 0.5rem', position: 'relative' }}>
                          <Upload size={20} style={{ margin: '0 auto 0.25rem', display: 'block', color: 'var(--color-text-muted)' }} />
                          <span style={{ fontSize: '0.75rem', fontWeight: 600 }}>Upload Trust Deed</span>
                          <input 
                            type="file" 
                            style={{ position: 'absolute', top: 0, left: 0, right: 0, bottom: 0, opacity: 0, cursor: 'pointer' }}
                            onChange={(e) => handleDocUpload(e, 'Trust_Deed')}
                          />
                        </div>
                      )}

                      {isEddRequired && (
                        <div className="file-uploader" style={{ padding: '1rem 0.5rem', position: 'relative' }}>
                          <Upload size={20} style={{ margin: '0 auto 0.25rem', display: 'block', color: 'var(--color-text-muted)' }} />
                          <span style={{ fontSize: '0.75rem', fontWeight: 600 }}>Upload SOW / Funds Proof</span>
                          <input 
                            type="file" 
                            style={{ position: 'absolute', top: 0, left: 0, right: 0, bottom: 0, opacity: 0, cursor: 'pointer' }}
                            onChange={(e) => handleDocUpload(e, 'EDD_Funds_Proof')}
                          />
                        </div>
                      )}
                    </div>

                    {isUploading && (
                      <div style={{ backgroundColor: 'var(--bg-app)', padding: '0.5rem', borderRadius: '4px' }}>
                        <div style={{ display: 'flex', justifySelf: 'space-between', fontSize: '0.7rem', fontWeight: 600, marginBottom: '2px' }}>
                          <span>Encrypting & Saving to SharePoint...</span>
                          <span>{uploadProgress}%</span>
                        </div>
                        <div style={{ height: '4px', width: '100%', backgroundColor: 'var(--color-border)', borderRadius: '2px', overflow: 'hidden' }}>
                          <div style={{ height: '100%', width: `${uploadProgress}%`, backgroundColor: 'var(--color-primary)', transition: 'width 0.1s' }} />
                        </div>
                      </div>
                    )}
                  </div>
                )}

                {/* Uploaded Documents List */}
                <div style={{ display: 'flex', flexDirection: 'column', gap: '0.35rem' }}>
                  <span style={{ fontSize: '0.75rem', fontWeight: 700, color: 'var(--color-text-secondary)' }}>SharePoint Foldered Attachments ({formData.documents.length})</span>
                  
                  {formData.documents.length > 0 ? (
                    formData.documents.map(doc => (
                      <div key={doc.id} className="uploaded-file-row" style={{ padding: '0.5rem 0.75rem' }}>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', overflow: 'hidden' }}>
                          <FileText size={16} style={{ color: 'var(--color-primary)', flexShrink: 0 }} />
                          <div style={{ display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
                            <span style={{ fontSize: '0.8rem', fontWeight: 600, textOverflow: 'ellipsis', overflow: 'hidden', whiteSpace: 'nowrap' }}>{doc.name}</span>
                            <span style={{ fontSize: '0.675rem', color: 'var(--color-text-muted)' }}>{doc.category} &bull; {doc.size}</span>
                          </div>
                        </div>
                        <span style={{ fontSize: '0.7rem', color: 'var(--color-low-risk)', fontWeight: 600 }}>SharePoint Verified</span>
                      </div>
                    ))
                  ) : (
                    <span style={{ fontSize: '0.75rem', color: 'var(--color-text-muted)', textAlign: 'center', padding: '1rem', display: 'block', backgroundColor: 'var(--bg-app)', border: '1px solid var(--color-border)', borderRadius: '6px' }}>
                      No files currently uploaded in SharePoint folder root.
                    </span>
                  )}
                </div>
                {errors.documents && <div style={{ color: 'var(--color-high-risk)', fontSize: '0.75rem', fontWeight: 600, marginTop: '0.5rem' }}>{errors.documents}</div>}
              </div>

            </div>
          </div>
        )}

        {/* Step 5: EDD Measures */}
        {activeStep === 5 && (
          <div style={{ display: 'flex', flexDirection: 'column', gap: '1.25rem' }}>
            <div className="alert-panel" style={{ borderLeftColor: 'var(--color-med-risk)', backgroundColor: 'var(--color-med-risk-bg)', color: '#78350f', margin: '0 0 0.5rem' }}>
              <ShieldCheck size={18} style={{ float: 'left', marginRight: '0.5rem' }} />
              <strong>Enhanced Due Diligence (EDD) Measures Enforced</strong>: Case portfolio meets medium/high risk criteria or PEP triggers. Source of Funds & Wealth establishment are statutory prerequisites.
            </div>

            <h3 style={{ fontSize: '1.1rem', fontWeight: 700, borderBottom: '2px solid var(--color-border)', paddingBottom: '0.5rem' }}>
              Step 5: Enhanced Due Diligence (EDD) – Medium & High Risk Cases
            </h3>

            <div className="grid-2">
              <div style={{ display: 'flex', flexDirection: 'column', gap: '0.85rem' }}>
                <span style={{ fontSize: '0.85rem', fontWeight: 700, color: 'var(--color-primary)' }}>Mandated EDD Measures Checklist</span>
                
                {[
                  { key: 'sourceOfFundsVerified', label: 'Source of funds established and verified with backing documents' },
                  { key: 'sourceOfWealthVerified', label: 'Source of wealth determined and verified via audited archives' },
                  { key: 'enhancedAdverseMedia', label: 'Enhanced/deep adverse media screening executed' },
                  { key: 'additionalBoVerification', label: 'Additional UBO identification verification' },
                  { key: 'enhancedMonitoringApplied', label: 'Enhanced and more frequent ongoing transaction monitoring applied' }
                ].map(item => (
                  <label key={item.key} style={{ display: 'flex', gap: '0.75rem', alignItems: 'center', cursor: isReadOnly ? 'not-allowed' : 'pointer', fontSize: '0.825rem' }}>
                    <input
                      type="checkbox"
                      checked={formData.eddApplied[item.key as keyof OnboardingCase['eddApplied']]}
                      disabled={isReadOnly}
                      onChange={(e) => {
                        const checked = e.target.checked;
                        setFormData(prev => ({
                          ...prev,
                          eddApplied: { ...prev.eddApplied, [item.key]: checked }
                        }));
                      }}
                    />
                    <span style={{ color: 'var(--color-text-secondary)', fontWeight: 500 }}>{item.label}</span>
                  </label>
                ))}
              </div>

              <div className="form-group">
                <label className="form-label">EDD Key Findings and Risk Mitigating Summaries <span className="required">*</span></label>
                <textarea 
                  name="eddFindings" 
                  rows={6} 
                  className="form-control" 
                  value={formData.eddFindings} 
                  onChange={handleChange} 
                  disabled={isReadOnly}
                  placeholder="Detail the source of wealth, source of funds, and state precise mitigating arrangements... (e.g. escrow deposits, audited funds path)"
                />
              </div>
            </div>
          </div>
        )}

        {/* Step 6: PEP / Screening */}
        {activeStep === 6 && (
          <div style={{ display: 'flex', flexDirection: 'column', gap: '1.5rem' }}>
            <h3 style={{ fontSize: '1.1rem', fontWeight: 700, borderBottom: '2px solid var(--color-border)', paddingBottom: '0.5rem' }}>
              Step 6: Politically Exposed Persons (PEP), Sanctions & Adverse News Screening
            </h3>

            <div className="grid-3" style={{ gap: '1.5rem' }}>
              
              {/* PEP Status Card */}
              <div style={{ padding: '1rem', backgroundColor: 'var(--bg-app)', border: '1px solid var(--color-border)', borderRadius: '8px' }}>
                <span style={{ fontSize: '0.85rem', fontWeight: 700, color: 'var(--color-primary)', display: 'block', marginBottom: '0.75rem' }}>PEP Screening Status</span>
                
                <div className="form-group">
                  <label className="form-label">Is Client / Beneficial Owner a PEP? <span className="required">*</span></label>
                  <div style={{ display: 'flex', gap: '1rem', marginTop: '0.25rem' }}>
                    {[
                      { val: false, lbl: 'No PEP Match' },
                      { val: true, lbl: 'PEP Hits Identified' }
                    ].map(opt => (
                      <label key={opt.lbl} style={{ display: 'flex', alignItems: 'center', gap: '0.25rem', fontSize: '0.8rem', cursor: isReadOnly ? 'not-allowed' : 'pointer' }}>
                        <input
                          type="radio"
                          name="isPep"
                          checked={formData.pepStatus.isPep === opt.val}
                          disabled={isReadOnly}
                          onChange={() => {
                            setFormData(prev => ({
                              ...prev,
                              pepStatus: { ...prev.pepStatus, isPep: opt.val, pepType: opt.val ? 'Domestic' : undefined }
                            }));
                          }}
                        />
                        {opt.lbl}
                      </label>
                    ))}
                  </div>
                </div>

                {formData.pepStatus.isPep && (
                  <div className="form-group" style={{ marginTop: '0.75rem' }}>
                    <label className="form-label">PEP Category Type</label>
                    <select
                      className="form-control"
                      style={{ fontSize: '0.8rem' }}
                      value={formData.pepStatus.pepType}
                      disabled={isReadOnly}
                      onChange={e => {
                        const val = e.target.value as any;
                        setFormData(prev => ({
                          ...prev,
                          pepStatus: { ...prev.pepStatus, pepType: val }
                        }));
                      }}
                    >
                      <option value="Domestic">Domestic PEP (Zimbabwe/Malawi Gov)</option>
                      <option value="Foreign">Foreign PEP (International Gov)</option>
                      <option value="International">International Organization Leader</option>
                    </select>
                  </div>
                )}
              </div>

              {/* Sanctions Screening Card */}
              <div style={{ padding: '1rem', backgroundColor: 'var(--bg-app)', border: '1px solid var(--color-border)', borderRadius: '8px' }}>
                <span style={{ fontSize: '0.85rem', fontWeight: 700, color: 'var(--color-primary)', display: 'block', marginBottom: '0.75rem' }}>UNSCR Sanctions Check</span>
                
                <label style={{ display: 'flex', gap: '0.5rem', alignItems: 'center', fontSize: '0.8rem', marginBottom: '0.75rem' }}>
                  <input
                    type="checkbox"
                    checked={formData.sanctionsCheck.screened}
                    disabled={isReadOnly}
                    onChange={(e) => {
                      const checked = e.target.checked;
                      setFormData(prev => ({
                        ...prev,
                        sanctionsCheck: { ...prev.sanctionsCheck, screened: checked }
                      }));
                    }}
                  />
                  <span>Screened against active UNSCR lists</span>
                </label>

                <div className="form-group">
                  <label className="form-label">Screening Match Result</label>
                  <div style={{ display: 'flex', gap: '1rem', marginTop: '0.25rem' }}>
                    {[
                      { val: false, lbl: 'Passed (No hit)' },
                      { val: true, lbl: 'Positive Hit Found' }
                    ].map(opt => (
                      <label key={opt.lbl} style={{ display: 'flex', alignItems: 'center', gap: '0.25rem', fontSize: '0.8rem', cursor: isReadOnly ? 'not-allowed' : 'pointer' }}>
                        <input
                          type="radio"
                          name="sanctionsMatch"
                          checked={formData.sanctionsCheck.hasMatch === opt.val}
                          disabled={isReadOnly}
                          onChange={() => {
                            setFormData(prev => ({
                              ...prev,
                              sanctionsCheck: { ...prev.sanctionsCheck, hasMatch: opt.val }
                            }));
                          }}
                        />
                        {opt.lbl}
                      </label>
                    ))}
                  </div>
                </div>

                {formData.sanctionsCheck.hasMatch && (
                  <textarea
                    rows={2}
                    className="form-control"
                    style={{ fontSize: '0.8rem', marginTop: '0.5rem' }}
                    placeholder="Enter positive match details or target identification matches..."
                    value={formData.sanctionsCheck.details || ''}
                    disabled={isReadOnly}
                    onChange={e => {
                      const val = e.target.value;
                      setFormData(prev => ({
                        ...prev,
                        sanctionsCheck: { ...prev.sanctionsCheck, details: val }
                      }));
                    }}
                  />
                )}
              </div>

              {/* Adverse Media Card */}
              <div style={{ padding: '1rem', backgroundColor: 'var(--bg-app)', border: '1px solid var(--color-border)', borderRadius: '8px' }}>
                <span style={{ fontSize: '0.85rem', fontWeight: 700, color: 'var(--color-primary)', display: 'block', marginBottom: '0.75rem' }}>Negative News & Adverse News</span>
                
                <div className="form-group">
                  <label className="form-label">Adverse News Identified?</label>
                  <div style={{ display: 'flex', gap: '1rem', marginTop: '0.25rem' }}>
                    {[
                      { val: false, lbl: 'No negative news' },
                      { val: true, lbl: 'Adverse news identified' }
                    ].map(opt => (
                      <label key={opt.lbl} style={{ display: 'flex', alignItems: 'center', gap: '0.25rem', fontSize: '0.8rem', cursor: isReadOnly ? 'not-allowed' : 'pointer' }}>
                        <input
                          type="radio"
                          name="adverseMedia"
                          checked={formData.adverseMediaCheck.hasAdverseInfo === opt.val}
                          disabled={isReadOnly}
                          onChange={() => {
                            setFormData(prev => ({
                              ...prev,
                              adverseMediaCheck: { ...prev.adverseMediaCheck, hasAdverseInfo: opt.val }
                            }));
                          }}
                        />
                        {opt.lbl}
                      </label>
                    ))}
                  </div>
                </div>

                {formData.adverseMediaCheck.hasAdverseInfo && (
                  <textarea
                    rows={3}
                    className="form-control"
                    style={{ fontSize: '0.8rem', marginTop: '0.5rem' }}
                    placeholder="Summarize adverse media findings and explain mitigating elements..."
                    value={formData.adverseMediaCheck.details || ''}
                    disabled={isReadOnly}
                    onChange={e => {
                      const val = e.target.value;
                      setFormData(prev => ({
                        ...prev,
                        adverseMediaCheck: { ...prev.adverseMediaCheck, details: val }
                      }));
                    }}
                  />
                )}
              </div>

            </div>
          </div>
        )}

        {/* Step 7: Declarations / Sign-off */}
        {activeStep === 7 && (
          <div style={{ display: 'flex', flexDirection: 'column', gap: '1.5rem' }}>
            <h3 style={{ fontSize: '1.1rem', fontWeight: 700, borderBottom: '2px solid var(--color-border)', paddingBottom: '0.5rem' }}>
              Step 7: Final Declarations & Onboarding Submission
            </h3>

            <div className="alert-panel" style={{ margin: 0, padding: '0.75rem', display: 'flex', gap: '0.5rem', alignItems: 'center' }}>
              <Info size={16} style={{ color: 'var(--color-primary)', flexShrink: 0 }} />
              <span style={{ fontSize: '0.75rem', lineHeight: '1.4' }}>
                On final submission, this onboarding item is uploaded directly to BDO Zimbabwe/Malawi SharePoint Lists. Power Automate triggers will locked editing access for the Preparer and notify Compliance Reviewers.
              </span>
            </div>

            <div className="grid-2">
              
              {/* Submission approvals status */}
              <div style={{ display: 'flex', flexDirection: 'column', gap: '1rem', padding: '1rem', backgroundColor: 'var(--bg-app)', borderRadius: '8px', border: '1px solid var(--color-border)' }}>
                <span style={{ fontSize: '0.85rem', fontWeight: 700, color: 'var(--color-primary)' }}>Preparer Submission Panel</span>
                
                <div className="form-group">
                  <label className="form-label">Compliance Declarations</label>
                  <div style={{ display: 'flex', gap: '0.5rem', fontSize: '0.8rem', color: 'var(--color-text-secondary)', marginTop: '0.25rem' }}>
                    <input type="checkbox" checked={true} readOnly disabled />
                    <span>I declare that all CDD documents have been verified in line with BDO Policies.</span>
                  </div>
                </div>

                <div style={{ display: 'flex', flexDirection: 'column', gap: '0.75rem', padding: '0.75rem', border: '1px dashed var(--color-border)', borderRadius: '6px' }}>
                  <span style={{ fontSize: '0.75rem', fontWeight: 700, color: 'var(--color-text-secondary)' }}>
                    Assign workflow reviewers (BDO email addresses)
                  </span>
                  <p style={{ fontSize: '0.7rem', color: 'var(--color-text-muted)', margin: 0, lineHeight: 1.4 }}>
                    Set who receives each approval step. On submit, the case routes to Compliance first, then Risk Partner (high risk or PEP), then Engagement Partner.
                  </p>

                  <div className="form-group" style={{ margin: 0 }}>
                    <label className="form-label">Compliance / MLRO reviewer <span className="required">*</span></label>
                    <input
                      type="email"
                      className="form-control"
                      placeholder="compliance.reviewer@bdo.co.zw"
                      value={formData.workflowAssignees.complianceEmail}
                      disabled={isReadOnly}
                      onChange={(e) =>
                        setFormData((prev) => ({
                          ...prev,
                          workflowAssignees: {
                            ...prev.workflowAssignees,
                            complianceEmail: e.target.value,
                          },
                        }))
                      }
                    />
                    {errors.complianceReviewerEmail && (
                      <span style={{ color: 'var(--color-high-risk)', fontSize: '0.75rem', fontWeight: 600 }}>
                        {errors.complianceReviewerEmail}
                      </span>
                    )}
                  </div>

                  <div className="form-group" style={{ margin: 0 }}>
                    <label className="form-label">Engagement Partner reviewer <span className="required">*</span></label>
                    <input
                      type="email"
                      className="form-control"
                      placeholder="engagement.partner@bdo.co.zw"
                      value={formData.workflowAssignees.engagementPartnerEmail}
                      disabled={isReadOnly}
                      onChange={(e) =>
                        setFormData((prev) => ({
                          ...prev,
                          workflowAssignees: {
                            ...prev.workflowAssignees,
                            engagementPartnerEmail: e.target.value,
                          },
                        }))
                      }
                    />
                    {errors.engagementPartnerReviewerEmail && (
                      <span style={{ color: 'var(--color-high-risk)', fontSize: '0.75rem', fontWeight: 600 }}>
                        {errors.engagementPartnerReviewerEmail}
                      </span>
                    )}
                  </div>

                  <div className="form-group" style={{ margin: 0 }}>
                    <label className="form-label">Risk Partner reviewer <span className="required">*</span></label>
                    <input
                      type="email"
                      className="form-control"
                      placeholder="risk.partner@bdo.co.zw"
                      value={formData.workflowAssignees.riskPartnerEmail}
                      disabled={isReadOnly}
                      onChange={(e) =>
                        setFormData((prev) => ({
                          ...prev,
                          workflowAssignees: {
                            ...prev.workflowAssignees,
                            riskPartnerEmail: e.target.value,
                          },
                        }))
                      }
                    />
                    {errors.riskPartnerReviewerEmail && (
                      <span style={{ color: 'var(--color-high-risk)', fontSize: '0.75rem', fontWeight: 600 }}>
                        {errors.riskPartnerReviewerEmail}
                      </span>
                    )}
                  </div>
                </div>

                <div className="form-group">
                  <label className="form-label">Electronic Signature / Sign-off Name <span className="required">*</span></label>
                  <input
                    type="text"
                    name="preparerSignature"
                    placeholder="Enter your full name as electronic signature..."
                    className="form-control"
                    style={{ fontWeight: 600, fontStyle: 'italic' }}
                    value={preparerSignature}
                    onChange={(e) => setPreparerSignature(e.target.value)}
                    disabled={isReadOnly}
                  />
                  {errors.signature && <span style={{ color: 'var(--color-high-risk)', fontSize: '0.75rem', fontWeight: 600 }}>{errors.signature}</span>}
                </div>
              </div>

              {/* Signature visualization pad */}
              <div style={{ display: 'flex', flexDirection: 'column', gap: '0.5rem', justifyContent: 'center' }}>
                <span style={{ fontSize: '0.75rem', fontWeight: 700, color: 'var(--color-text-secondary)' }}>Corporate Signatures Timeline Summary</span>
                
                <div style={{ padding: '1rem', backgroundColor: 'var(--bg-card)', border: '1px solid var(--color-border)', borderRadius: '8px', display: 'flex', flexDirection: 'column', gap: '0.5rem', fontSize: '0.8rem' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                    <span><strong>1. Prepared by (Team):</strong></span>
                    <span style={{ color: formData.signatures.preparer ? 'var(--color-low-risk)' : 'var(--color-text-muted)', fontWeight: 600 }}>
                      {formData.signatures.preparer ? `Signed: ${formData.signatures.preparer.sign}` : 'Pending Submission'}
                    </span>
                  </div>
                  
                  <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                    <span><strong>2. Compliance MLRO Review:</strong></span>
                    <span style={{ color: formData.signatures.compliance ? 'var(--color-low-risk)' : 'var(--color-text-muted)', fontWeight: 600 }}>
                      {formData.signatures.compliance ? `Signed: ${formData.signatures.compliance.sign}` : 'Pending Compliance'}
                    </span>
                  </div>

                  <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                    <span><strong>3. Approved by Risk Partner (High Risk):</strong></span>
                    <span style={{ color: formData.signatures.riskPartner ? 'var(--color-low-risk)' : 'var(--color-text-muted)', fontWeight: 600 }}>
                      {formData.overallRiskRating === 'High' ? (
                        formData.signatures.riskPartner ? `Signed: ${formData.signatures.riskPartner.sign}` : 'Pending Risk Sign-off'
                      ) : 'Not Triggered (Low/Med Risk)'}
                    </span>
                  </div>

                  <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                    <span><strong>4. Approved by Engagement Partner:</strong></span>
                    <span style={{ color: formData.signatures.engagementPartner ? 'var(--color-low-risk)' : 'var(--color-text-muted)', fontWeight: 600 }}>
                      {formData.signatures.engagementPartner ? `Signed: ${formData.signatures.engagementPartner.sign}` : 'Pending Partner Approval'}
                    </span>
                  </div>
                </div>
              </div>

            </div>
          </div>
        )}

      </div>

      {/* Stepper Buttons Panel */}
      <div className="flex-between">
        <button 
          className="btn btn-secondary" 
          disabled={activeStep === 1}
          onClick={() => {
            // Skip EDD if not required when moving back
            if (activeStep === 6 && !isEddRequired) setActiveStep(4);
            else setActiveStep(prev => prev - 1);
          }}
        >
          <ChevronLeft size={16} />
          Back Component
        </button>

        <div style={{ display: 'flex', gap: '0.75rem' }}>
          
          {/* Save Draft Button */}
          {!isReadOnly && (
            <button className="btn btn-secondary" onClick={handleSaveDraft}>
              <Save size={16} />
              Save draft
            </button>
          )}

          {activeStep < 7 ? (
            <button 
              className="btn btn-primary"
              onClick={() => {
                // Skip EDD step if not required
                if (activeStep === 4 && !isEddRequired) setActiveStep(6);
                else setActiveStep(prev => prev + 1);
              }}
            >
              Continue
              <ChevronRight size={16} />
            </button>
          ) : (
            !isReadOnly && (
              <button className="btn btn-success" onClick={handleFinalSubmit}>
                <Send size={16} />
                Submit to Compliance
              </button>
            )
          )}
        </div>
      </div>

    </div>
  );
};
