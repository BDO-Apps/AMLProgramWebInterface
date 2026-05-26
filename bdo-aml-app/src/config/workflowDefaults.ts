/** Default reviewer emails for AML workflow routing (BDO test personas) */
export interface WorkflowAssignees {
  complianceEmail: string;
  engagementPartnerEmail: string;
  riskPartnerEmail: string;
}

export const defaultWorkflowAssignees: WorkflowAssignees = {
  complianceEmail: 'tmutasa@bdo.co.zw',
  engagementPartnerEmail: 'cmariro@bdo.co.zw',
  riskPartnerEmail: 'amugumwa@bdo.co.zw',
};

export function getDefaultWorkflowAssignees(): WorkflowAssignees {
  return { ...defaultWorkflowAssignees };
}

export function isValidEmail(email: string): boolean {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email.trim());
}
