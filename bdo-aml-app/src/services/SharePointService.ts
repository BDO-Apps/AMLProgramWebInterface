/**
 * SharePoint service facade — uses Microsoft Graph against BDO AuditSoftwareDevSite,
 * with optional localStorage mock when VITE_USE_SHAREPOINT_MOCK=true.
 */

import { sharePointConfig } from '../config/sharepointConfig';
import {
  getDefaultWorkflowAssignees,
  type WorkflowAssignees,
} from '../config/workflowDefaults';
import { sharePointRepository } from './sharepoint/SharePointRepository';

export type { WorkflowAssignees };
export { getDefaultWorkflowAssignees };

export interface Director {
  id?: string;
  fullName: string;
  position: string;
  nationality: string;
  idNumber: string;
  countryOfResidence: string;
}

export interface BeneficialOwner {
  id?: string;
  fullName: string;
  ownershipPercentage: number;
  basisOfControl: string;
  country: string;
  verificationSource: string;
}

export interface DocumentUpload {
  id: string;
  name: string;
  category: 'ID_Passport' | 'Incorporation_Doc' | 'CR6_Directors' | 'Trust_Deed' | 'BO_Declaration' | 'EDD_Funds_Proof' | 'PEP_Mitigation';
  size: string;
  uploadDate: string;
  uploadedBy: string;
}

export interface AuditLog {
  id: string;
  timestamp: string;
  actor: string;
  role: string;
  action: string;
  comments: string;
}

export interface OnboardingCase {
  id: string;
  sharePointItemId?: string;
  clientName: string;
  clientType: 'Individual' | 'Legal Person' | 'Legal Arrangement';
  regNumber: string;
  registeredAddress: string;
  natureOfBusiness: string;
  purposeOfEngagement: string;
  contactInfo: {
    address: string;
    email: string;
    phone: string;
  };
  directors: Director[];
  beneficialOwners: BeneficialOwner[];
  hasBeneficialOwners: boolean;
  riskRatings: {
    client: 'Low' | 'Medium' | 'High';
    geography: 'Low' | 'Medium' | 'High';
    productService: 'Low' | 'Medium' | 'High';
    deliveryChannel: 'Low' | 'Medium' | 'High';
    paymentMode: 'Low' | 'Medium' | 'High';
  };
  overallRiskRating: 'Low' | 'Medium' | 'High';
  riskRationale: string;
  cddMeasures: {
    identityVerified: boolean;
    boVerified: boolean;
    natureUnderstood: boolean;
    pepScreened: boolean;
    sanctionsScreened: boolean;
    adverseMediaScreened: boolean;
  };
  eddApplied: {
    sourceOfFundsVerified: boolean;
    sourceOfWealthVerified: boolean;
    enhancedAdverseMedia: boolean;
    additionalBoVerification: boolean;
    seniorMgmtApproved: boolean;
    enhancedMonitoringApplied: boolean;
  };
  eddFindings: string;
  pepStatus: {
    isPep: boolean;
    pepType?: 'Domestic' | 'Foreign' | 'International';
  };
  sanctionsCheck: {
    screened: boolean;
    hasMatch: boolean;
    details?: string;
  };
  adverseMediaCheck: {
    hasAdverseInfo: boolean;
    details?: string;
  };
  decision: 'Client accepted' | 'Client accepted subject to conditions/EDD' | 'Client declined' | '';
  reviewFrequency: 'Periodic' | 'Annual' | 'Enhanced and continuous' | '';
  nextReviewDate: string;
  signatures: {
    preparer?: { name: string; date: string; sign: string };
    compliance?: { name: string; date: string; sign: string };
    engagementPartner?: { name: string; date: string; sign: string };
    riskPartner?: { name: string; date: string; sign: string };
  };
  status: 'Draft' | 'Pending Compliance' | 'Pending Engagement Partner' | 'Pending Risk Partner' | 'Approved' | 'Returned' | 'Rejected';
  currentHandler: string;
  /** Reviewer emails assigned by preparer at sign-off (Step 7) */
  workflowAssignees: WorkflowAssignees;
  documents: DocumentUpload[];
  auditLogs: AuditLog[];
  office: 'Zimbabwe' | 'Malawi';
  dateCreated: string;
  lastUpdated: string;
}

const STORAGE_KEY = 'bdo_aml_cases_db';

/** In-memory mock used only when VITE_USE_SHAREPOINT_MOCK=true */
class MockSharePointStore {
  private cases: OnboardingCase[] = [];

  constructor() {
    this.loadFromStorage();
  }

  private loadFromStorage() {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) {
      try {
        this.cases = JSON.parse(saved);
        return;
      } catch {
        this.cases = [];
      }
    }
    this.seedInitialData();
  }

  private saveToStorage() {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(this.cases));
  }

  public clearStorage() {
    localStorage.removeItem(STORAGE_KEY);
    this.seedInitialData();
  }

  private seedInitialData() {
    this.cases = [
      {
        id: 'BDO-AML-2026-0001',
        clientName: 'Kariba Minerals Pvt Ltd',
        clientType: 'Legal Person',
        regNumber: 'CO-7782/2019',
        registeredAddress: '12 Fife Avenue, Bulawayo, Zimbabwe',
        natureOfBusiness: 'Mining and Export of Gemstones',
        purposeOfEngagement: 'Annual Statutory Audit and Tax Compliance Services',
        contactInfo: {
          address: '12 Fife Avenue, Bulawayo',
          email: 'finance@karibaminerals.co.zw',
          phone: '+263 9 881234',
        },
        directors: [
          {
            fullName: 'Simbarashe Moyo',
            position: 'Managing Director',
            nationality: 'Zimbabwean',
            idNumber: '63-123456-A-45',
            countryOfResidence: 'Zimbabwe',
          },
        ],
        beneficialOwners: [],
        hasBeneficialOwners: true,
        riskRatings: {
          client: 'High',
          geography: 'Medium',
          productService: 'Medium',
          deliveryChannel: 'Low',
          paymentMode: 'High',
        },
        overallRiskRating: 'High',
        riskRationale: 'High-risk extractive sector client.',
        cddMeasures: {
          identityVerified: true,
          boVerified: true,
          natureUnderstood: true,
          pepScreened: true,
          sanctionsScreened: true,
          adverseMediaScreened: true,
        },
        eddApplied: {
          sourceOfFundsVerified: true,
          sourceOfWealthVerified: true,
          enhancedAdverseMedia: true,
          additionalBoVerification: true,
          seniorMgmtApproved: false,
          enhancedMonitoringApplied: true,
        },
        eddFindings: '',
        pepStatus: { isPep: false },
        sanctionsCheck: { screened: true, hasMatch: false },
        adverseMediaCheck: { hasAdverseInfo: false },
        decision: '',
        reviewFrequency: '',
        nextReviewDate: '',
        signatures: {},
        status: 'Draft',
        currentHandler: 't.moyo@bdo.co.zw',
        workflowAssignees: getDefaultWorkflowAssignees(),
        documents: [],
        auditLogs: [],
        office: 'Zimbabwe',
        dateCreated: '2026-05-18',
        lastUpdated: '2026-05-18',
      },
    ];
    this.saveToStorage();
  }

  async getCases(role: string, email: string) {
    if (role === 'Preparer') {
      return this.cases.filter(
        (c) =>
          c.auditLogs.some((log) => log.actor === email) ||
          c.currentHandler === email ||
          c.status === 'Draft'
      );
    }
    if (role === 'RiskPartner') {
      return this.cases.filter((c) => c.overallRiskRating === 'High');
    }
    return [...this.cases];
  }

  async getCaseById(id: string) {
    const found = this.cases.find((c) => c.id === id);
    return found ? JSON.parse(JSON.stringify(found)) : null;
  }

  async saveCase(caseData: OnboardingCase, _actor: string, _role: string) {
    const cleanCase = { ...caseData, lastUpdated: new Date().toISOString().split('T')[0] };
    const idx = this.cases.findIndex((c) => c.id === cleanCase.id);
    if (idx >= 0) {
      this.cases[idx] = cleanCase;
    } else {
      cleanCase.id = `BDO-AML-${new Date().getFullYear()}-${String(this.cases.length + 1).padStart(4, '0')}`;
      cleanCase.dateCreated = cleanCase.lastUpdated;
      this.cases.push(cleanCase);
    }
    this.saveToStorage();
    return cleanCase;
  }

  async submitCase(
    caseId: string,
    actor: string,
    _role: string,
    preparerSign: string,
    assignees?: WorkflowAssignees
  ) {
    const c = this.cases.find((x) => x.id === caseId);
    if (!c) throw new Error('Case not found');
    if (!c.workflowAssignees) c.workflowAssignees = getDefaultWorkflowAssignees();
    if (assignees) c.workflowAssignees = assignees;
    c.status = 'Pending Compliance';
    c.currentHandler = c.workflowAssignees.complianceEmail.toLowerCase();
    c.signatures.preparer = { name: actor, date: new Date().toISOString().split('T')[0], sign: preparerSign };
    this.saveToStorage();
    return JSON.parse(JSON.stringify(c));
  }

  async processWorkflowAction(
    caseId: string,
    action: 'Approve' | 'Return' | 'Reject',
    comments: string,
    actor: string,
    role: string,
    signatureSign: string
  ) {
    const c = this.cases.find((x) => x.id === caseId);
    if (!c) throw new Error('Case not found');

    const timestamp = new Date().toISOString().split('T')[0];
    let nextStatus = c.status;
    let nextHandler = c.currentHandler;

    if (action === 'Reject') {
      nextStatus = 'Rejected';
      nextHandler = '';
      c.decision = 'Client declined';
    } else if (action === 'Return') {
      nextStatus = 'Returned';
      nextHandler = c.auditLogs.find((l) => l.action === 'CREATED')?.actor ?? '';
    } else if (role === 'Compliance') {
      c.signatures.compliance = {
        name: actor,
        date: timestamp,
        sign: signatureSign,
      };
      if (c.overallRiskRating === 'High' || c.pepStatus.isPep) {
        nextStatus = 'Pending Risk Partner';
        nextHandler = c.workflowAssignees.riskPartnerEmail.toLowerCase();
      } else {
        nextStatus = 'Pending Engagement Partner';
        nextHandler = c.workflowAssignees.engagementPartnerEmail.toLowerCase();
      }
    } else if (role === 'RiskPartner') {
      c.signatures.riskPartner = {
        name: actor,
        date: timestamp,
        sign: signatureSign,
      };
      nextStatus = 'Pending Engagement Partner';
      nextHandler = c.workflowAssignees.engagementPartnerEmail.toLowerCase();
    } else if (role === 'EngagementPartner') {
      c.signatures.engagementPartner = {
        name: actor,
        date: timestamp,
        sign: signatureSign,
      };
      nextStatus = 'Approved';
      nextHandler = '';

      if (c.overallRiskRating === 'High' || c.pepStatus.isPep) {
        c.decision = 'Client accepted subject to conditions/EDD';
      } else {
        c.decision = 'Client accepted';
      }

      const baseDate = new Date();
      if (c.overallRiskRating === 'High') {
        baseDate.setMonth(baseDate.getMonth() + 6);
        c.reviewFrequency = 'Enhanced and continuous';
      } else if (c.overallRiskRating === 'Medium') {
        baseDate.setFullYear(baseDate.getFullYear() + 1);
        c.reviewFrequency = 'Annual';
      } else {
        baseDate.setFullYear(baseDate.getFullYear() + 2);
        c.reviewFrequency = 'Periodic';
      }
      c.nextReviewDate = baseDate.toISOString().split('T')[0];
    }

    c.status = nextStatus;
    c.currentHandler = nextHandler;
    c.lastUpdated = timestamp;

    c.auditLogs.push({
      id: `log-${Date.now()}`,
      timestamp: new Date().toISOString(),
      actor,
      role,
      action: action.toUpperCase(),
      comments: comments || `${role} executed ${action.toLowerCase()} sign-off.`,
    });

    this.saveToStorage();
    return JSON.parse(JSON.stringify(c));
  }

  async uploadDocument(
    caseId: string,
    file: File,
    category: DocumentUpload['category'],
    actor: string
  ) {
    const c = this.cases.find((x) => x.id === caseId);
    if (!c) throw new Error('Case not found');
    const doc: DocumentUpload = {
      id: 'doc-' + Date.now(),
      name: file.name,
      category,
      size: `${(file.size / 1024 / 1024).toFixed(1)} MB`,
      uploadDate: new Date().toISOString().split('T')[0],
      uploadedBy: actor,
    };
    c.documents.push(doc);
    this.saveToStorage();
    return doc;
  }
}

class SharePointService {
  public readonly isDemoMode = sharePointConfig.useMock;
  private mock = new MockSharePointStore();

  public async getCases(role: string, email: string): Promise<OnboardingCase[]> {
    if (this.isDemoMode) return this.mock.getCases(role, email);
    return sharePointRepository.getCases(role, email);
  }

  public async getCaseById(id: string): Promise<OnboardingCase | null> {
    if (this.isDemoMode) return this.mock.getCaseById(id);
    return sharePointRepository.getCaseById(id);
  }

  public async saveCase(
    caseData: OnboardingCase,
    actor: string,
    role: string
  ): Promise<OnboardingCase> {
    if (this.isDemoMode) return this.mock.saveCase(caseData, actor, role);
    return sharePointRepository.saveCase(caseData, actor, role);
  }

  public async submitCase(
    caseId: string,
    actor: string,
    role: string,
    preparerSign: string,
    assignees?: WorkflowAssignees
  ): Promise<OnboardingCase> {
    if (this.isDemoMode) return this.mock.submitCase(caseId, actor, role, preparerSign, assignees);
    return sharePointRepository.submitCase(caseId, actor, role, preparerSign, assignees);
  }

  public async processWorkflowAction(
    caseId: string,
    action: 'Approve' | 'Return' | 'Reject',
    comments: string,
    actor: string,
    role: string,
    signatureSign: string
  ): Promise<OnboardingCase> {
    if (this.isDemoMode) {
      return this.mock.processWorkflowAction(
        caseId,
        action,
        comments,
        actor,
        role,
        signatureSign
      );
    }
    return sharePointRepository.processWorkflowAction(
      caseId,
      action,
      comments,
      actor,
      role,
      signatureSign
    );
  }

  public async uploadDocument(
    caseId: string,
    file: File,
    category: DocumentUpload['category'],
    actor: string
  ): Promise<DocumentUpload> {
    if (this.isDemoMode) return this.mock.uploadDocument(caseId, file, category, actor);
    return sharePointRepository.uploadDocument(caseId, file, category, actor);
  }

  public async diagnoseColumns(): Promise<Array<{ name: string; displayName: string; type: string }>> {
    if (this.isDemoMode) return [];
    return sharePointRepository.diagnoseColumns();
  }

  public clearStorage() {
    if (!this.isDemoMode) {
      console.warn(
        'clearStorage() only applies in mock mode. Set VITE_USE_SHAREPOINT_MOCK=true to use local demo data.'
      );
      return;
    }
    this.mock.clearStorage();
  }
}

export const sharePointService = new SharePointService();
