import { sharePointConfig } from '../../config/sharepointConfig';
import { graphClient } from './graphClient';

export interface ListColumnInfo {
  name: string;
  displayName: string;
  type: string;
  readOnly: boolean;
}

/** SharePoint system / computed columns that must never be written via Graph */
const NON_WRITABLE_COLUMN_NAMES = new Set([
  'linktitle',
  'id',
  'author',
  'editor',
  'created',
  'modified',
  'contenttype',
  'attachments',
  'edit',
  'itemchildcount',
  'folderchildcount',
  'appauthor',
  'appeditor',
  '_uiversionstring',
  'complianceassetid',
]);

/** Logical app field → possible SharePoint internal names (first match wins) */
const CLIENT_ONBOARDING_ALIASES: Record<string, readonly string[]> = {
  Title: ['Title'],
  ClientName: ['ClientName', 'Client_x0020_Name'],
  ClientType: ['ClientType', 'Client_x0020_Type'],
  RegNumber: ['RegNumber', 'Registration_x0020_Number'],
  RegisteredAddress: ['RegisteredAddress', 'Registered_x0020_Address'],
  NatureOfBusiness: ['NatureOfBusiness', 'Nature_x0020_of_x0020_Business'],
  PurposeOfEngagement: ['PurposeOfEngagement', 'Purpose_x0020_of_x0020_Engagement'],
  ContactAddress: ['ContactAddress', 'Contact_x0020_Address'],
  ContactEmail: ['ContactEmail', 'Contact_x0020_Email', 'Contact_Email'],
  ContactPhone: ['ContactPhone', 'Contact_x0020_Phone', 'Contact_Phone'],
  ContactInfo: ['ContactInfo', 'Contact_x0020_Info'],
  Office: ['Office'],
  HasBeneficialOwners: ['HasBeneficialOwners', 'Has_x0020_Beneficial_x0020_Owners'],
  RiskRatingClient: ['RiskRatingClient', 'Risk_x0020_client', 'Risk_x0020_Client'],
  RiskRatingGeography: ['RiskRatingGeography', 'Risk_x0020_geography', 'Risk_x0020_Geography'],
  RiskRatingProductService: [
    'RiskRatingProductService',
    'Risk_x0020_product_x002f_service',
    'Risk_x0020_Product_x0020_Service',
  ],
  RiskRatingDeliveryChannel: [
    'RiskRatingDeliveryChannel',
    'Risk_x0020_delivery_x0020_channel',
    'Risk_x0020_Delivery_x0020_Channel',
  ],
  RiskRatingPaymentMode: [
    'RiskRatingPaymentMode',
    'Risk_x0020_payment_x0020_mode',
    'Risk_x0020_Payment_x0020_Mode',
  ],
  OverallRiskRating: ['OverallRiskRating', 'Overall_x0020_Risk_x0020_Rating'],
  RiskRationale: ['RiskRationale', 'Risk_x0020_Rationale'],
  // Section 3: Risk indicator checklist (flat columns approach)
  ClientRiskPEPOrAssociate: ['ClientRiskPEPOrAssociate'],
  ClientRiskNPO: ['ClientRiskNPO'],
  ClientRiskComplexOwnership: ['ClientRiskComplexOwnership'],
  ClientRiskCashIntensive: ['ClientRiskCashIntensive'],
  ClientRiskIntermediaries: ['ClientRiskIntermediaries'],
  ClientRiskOther: ['ClientRiskOther'],
  GeoRiskFATFGreyListed: ['GeoRiskFATFGreyListed'],
  GeoRiskFATFBlackListed: ['GeoRiskFATFBlackListed'],
  GeoRiskSanctionsExposed: ['GeoRiskSanctionsExposed'],
  GeoRiskHighCorruptionConflict: ['GeoRiskHighCorruptionConflict'],
  GeoRiskOther: ['GeoRiskOther'],
  ProdRiskTrustCompanyFormation: ['ProdRiskTrustCompanyFormation'],
  ProdRiskManageClientFundsAssets: ['ProdRiskManageClientFundsAssets'],
  ProdRiskCrossBorderTransactions: ['ProdRiskCrossBorderTransactions'],
  ProdRiskHighValueComplexTransactions: ['ProdRiskHighValueComplexTransactions'],
  ProdRiskOther: ['ProdRiskOther'],
  CDDIdentityVerified: ['CDDIdentityVerified'],
  CDDBOVerified: ['CDDBOVerified'],
  CDDNatureUnderstood: ['CDDNatureUnderstood'],
  CDDPEPScreened: ['CDDPEPScreened'],
  CDDSanctionsScreened: ['CDDSanctionsScreened'],
  CDDAdverseMediaScreened: ['CDDAdverseMediaScreened'],
  EDDSourceOfFundsVerified: ['EDDSourceOfFundsVerified'],
  EDDSourceOfWealthVerified: ['EDDSourceOfWealthVerified'],
  EDDEnhancedAdverseMedia: ['EDDEnhancedAdverseMedia'],
  EDDAdditionalBOVerification: ['EDDAdditionalBOVerification'],
  EDDSeniorMgmtApproved: ['EDDSeniorMgmtApproved'],
  EDDEnhancedMonitoringApplied: ['EDDEnhancedMonitoringApplied'],
  EDDFindings: ['EDDFindings'],
  IsPEP: ['IsPEP', 'Is_x0020_PEP'],
  PEPType: ['PEPType', 'PEP_x0020_Type'],
  SanctionsScreened: ['SanctionsScreened'],
  SanctionsHasMatch: ['SanctionsHasMatch'],
  SanctionsDetails: ['SanctionsDetails'],
  AdverseMediaHasInfo: ['AdverseMediaHasInfo'],
  AdverseMediaDetails: ['AdverseMediaDetails'],
  Decision: ['Decision'],
  ReviewFrequency: ['ReviewFrequency', 'Review_x0020_Frequency'],
  NextReviewDate: ['NextReviewDate', 'Next_x0020_Review_x0020_Date'],
  PreparerSignName: ['PreparerSignName'],
  PreparerSignDate: ['PreparerSignDate'],
  PreparerSignText: ['PreparerSignText'],
  ComplianceSignName: ['ComplianceSignName'],
  ComplianceSignDate: ['ComplianceSignDate'],
  ComplianceSignText: ['ComplianceSignText'],
  EngagementPartnerSignName: ['EngagementPartnerSignName'],
  EngagementPartnerSignDate: ['EngagementPartnerSignDate'],
  EngagementPartnerSignText: ['EngagementPartnerSignText'],
  RiskPartnerSignName: ['RiskPartnerSignName'],
  RiskPartnerSignDate: ['RiskPartnerSignDate'],
  RiskPartnerSignText: ['RiskPartnerSignText'],
  ComplianceReviewComment: ['ComplianceReviewComment'],
  EPReviewComment: ['EPReviewComment'],
  RPReviewComment: ['RPReviewComment'],
  WorkflowStatus: ['WorkflowStatus', 'Workflow_x0020_Status'],
  CurrentHandler: ['CurrentHandler', 'Current_x0020_Handler'],
  ComplianceReviewerEmail: [
    'ComplianceReviewerEmail',
    'Compliance_x0020_Reviewer_x0020_Email',
    'ComplianceEmail',
  ],
  EngagementPartnerReviewerEmail: [
    'EngagementPartnerReviewerEmail',
    'EngagementPartnerEmail',
    'Engagement_x0020_Partner_x0020_Email',
  ],
  RiskPartnerReviewerEmail: [
    'RiskPartnerReviewerEmail',
    'RiskPartnerEmail',
    'Risk_x0020_Partner_x0020_Email',
  ],
  /** Optional JSON fallback when flat reviewer email columns are missing */
  WorkflowAssigneesJson: ['WorkflowAssigneesJson', 'WorkflowAssignees'],
  DateCreated: ['DateCreated', 'Date_x0020_Created', 'Created'],
  LastUpdated: ['LastUpdated', 'Last_x0020_Updated', 'Modified'],
};

const DIRECTORS_ALIASES: Record<string, readonly string[]> = {
  Title: ['Title'],
  FullName: ['FullName', 'Full_x0020_Name'],
  Position: ['Position'],
  Nationality: ['Nationality'],
  IDNumber: ['IDNumber', 'ID_x0020_Number'],
  CountryOfResidence: ['CountryOfResidence', 'Country_x0020_of_x0020_Residence'],
  CaseIDLookupId: ['CaseIDLookupId', 'CaseID_x003a_IDLookupId', 'CaseID'],
};

const BO_ALIASES: Record<string, readonly string[]> = {
  Title: ['Title'],
  FullName: ['FullName'],
  OwnershipPercentage: ['OwnershipPercentage', 'Ownership_x0020_Percentage'],
  BasisOfControl: ['BasisOfControl', 'Basis_x0020_of_x0020_Control'],
  Country: ['Country'],
  VerificationSource: ['VerificationSource', 'Verification_x0020_Source'],
  CaseIDLookupId: ['CaseIDLookupId', 'CaseID_x003a_IDLookupId', 'CaseID'],
};

const AUDIT_ALIASES: Record<string, readonly string[]> = {
  EventTimestamp: ['EventTimestamp', 'Event_x0020_Timestamp', 'Timestamp'],
  Actor: ['Actor'],
  Role: ['Role'],
  Action: ['Action'],
  Comments: ['Comments'],
  CaseIDLookupId: ['CaseIDLookupId', 'CaseID_x003a_IDLookupId', 'CaseID'],
};

const DOCUMENT_ALIASES: Record<string, readonly string[]> = {
  CaseID: ['CaseID', 'Case_x0020_ID'],
  DocumentCategory: ['DocumentCategory', 'Document_x0020_Category'],
  UploadedBy: ['UploadedBy', 'Uploaded_x0020_By'],
  UploadDate: ['UploadDate', 'Upload_x0020_Date'],
};

type AliasMap = Record<string, readonly string[]>;

class ColumnSchemaCache {
  private columnsByList = new Map<string, ListColumnInfo[]>();
  private nameIndexByList = new Map<string, Map<string, string>>();
  private writableByList = new Map<string, Set<string>>();
  private lookupIdFieldByList = new Map<string, string>();

  private normalize(key: string): string {
    return key.replace(/[_\s]/g, '').toLowerCase();
  }

  async loadColumns(siteId: string, listId: string): Promise<ListColumnInfo[]> {
    const cacheKey = `${siteId}:${listId}`;
    const cached = this.columnsByList.get(cacheKey);
    if (cached) return cached;

    const data = await graphClient.getListColumns(siteId, listId);
    this.columnsByList.set(cacheKey, data);

    const index = new Map<string, string>();
    const writable = new Set<string>();

    for (const col of data) {
      index.set(this.normalize(col.name), col.name);
      index.set(this.normalize(col.displayName), col.name);

      if (!col.readOnly && !NON_WRITABLE_COLUMN_NAMES.has(this.normalize(col.name))) {
        writable.add(col.name);
      }
    }
    this.nameIndexByList.set(cacheKey, index);
    this.writableByList.set(cacheKey, writable);
    this.lookupIdFieldByList.set(cacheKey, this.detectLookupIdFieldName(data));

    console.info(
      `[SharePoint] Loaded ${data.length} columns for list ${listId}:`,
      data.map((c) => c.name).join(', ')
    );

    return data;
  }

  private detectLookupIdFieldName(columns: ListColumnInfo[]): string {
    const base = sharePointConfig.caseLookupField;

    const lookupCol = columns.find(
      (c) =>
        c.type === 'lookup' &&
        (c.name.toLowerCase() === base.toLowerCase() ||
          c.displayName.toLowerCase().replace(/\s/g, '') === base.toLowerCase())
    );
    if (lookupCol) {
      return `${lookupCol.name}LookupId`;
    }

    const encoded = columns.find((c) =>
      c.name.toLowerCase().includes('caseid') && c.name.toLowerCase().includes('lookupid')
    );
    if (encoded) return encoded.name;

    return `${base}LookupId`;
  }

  getLookupIdFieldName(siteId: string, listId: string): string {
    const cacheKey = `${siteId}:${listId}`;
    return this.lookupIdFieldByList.get(cacheKey) ?? `${sharePointConfig.caseLookupField}LookupId`;
  }

  private resolveOne(
    cacheKey: string,
    aliases: readonly string[]
  ): string | null {
    const index = this.nameIndexByList.get(cacheKey);
    if (!index) return null;

    for (const alias of aliases) {
      const hit = index.get(this.normalize(alias));
      if (hit) return hit;
    }
    return null;
  }

  /**
   * Maps logical field keys to actual SharePoint column names.
   * Drops fields with no matching column. Omits empty optional values.
   */
  async mapFields(
    siteId: string,
    listId: string,
    logicalFields: Record<string, unknown>,
    aliasMap: AliasMap
  ): Promise<Record<string, unknown>> {
    await this.loadColumns(siteId, listId);
    const cacheKey = `${siteId}:${listId}`;
    const result: Record<string, unknown> = {};

    const lookupIdFieldName = this.getLookupIdFieldName(siteId, listId);

    for (const [logicalKey, value] of Object.entries(logicalFields)) {
      if (value === undefined) continue;
      if (value === null || value === '') continue;

      // Lookup IDs are set via {LookupColumn}LookupId — not a browsable column in schema
      if (logicalKey.endsWith('LookupId') || logicalKey === 'CaseIDLookupId') {
        result[lookupIdFieldName] = value;
        continue;
      }

      const aliases = aliasMap[logicalKey] ?? [logicalKey];
      const spName = this.resolveOne(cacheKey, aliases);
      if (!spName) continue;

      // Never write to the lookup display column directly (use LookupId field above)
      if (spName === sharePointConfig.caseLookupField) {
        continue;
      }

      const writable = this.writableByList.get(cacheKey);
      if (!writable?.has(spName)) {
        continue;
      }

      result[spName] = value;
    }

    return result;
  }

  /** Read a field from SharePoint item using alias list */
  getFieldValue(
    fields: Record<string, unknown>,
    aliases: readonly string[]
  ): unknown {
    for (const alias of aliases) {
      if (alias in fields) return fields[alias];
    }
    const normalizedFields = new Map(
      Object.entries(fields).map(([k, v]) => [this.normalize(k), v])
    );
    for (const alias of aliases) {
      const hit = normalizedFields.get(this.normalize(alias));
      if (hit !== undefined) return hit;
    }
    return undefined;
  }

  clearCache() {
    this.columnsByList.clear();
    this.nameIndexByList.clear();
    this.writableByList.clear();
    this.lookupIdFieldByList.clear();
  }
}

export const columnSchema = new ColumnSchemaCache();

export const aliasMaps = {
  clientOnboarding: CLIENT_ONBOARDING_ALIASES,
  directors: DIRECTORS_ALIASES,
  beneficialOwners: BO_ALIASES,
  auditLogs: AUDIT_ALIASES,
  documents: DOCUMENT_ALIASES,
};
