import { jsPDF } from 'jspdf';
import type { OnboardingCase } from '../services/SharePointService';

function checkbox(v: boolean): string {
  return v ? '[x]' : '[ ]';
}

function safe(v: string | undefined | null): string {
  return (v ?? '').trim();
}

function splitText(doc: jsPDF, text: string, maxWidth: number): string[] {
  const t = text || '';
  return doc.splitTextToSize(t, maxWidth) as string[];
}

export function downloadCasePdfExtract(caseData: OnboardingCase) {
  const doc = new jsPDF({ unit: 'pt', format: 'a4' });
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();

  const marginX = 40;
  const marginY = 44;
  const maxWidth = pageWidth - marginX * 2;

  let y = marginY;

  const line = (text: string, size = 10, extraY = 14, style: 'normal' | 'bold' = 'normal') => {
    doc.setFont('helvetica', style);
    doc.setFontSize(size);
    doc.text(text, marginX, y, { maxWidth });
    y += extraY;
  };

  const hr = () => {
    doc.setDrawColor(180);
    doc.line(marginX, y, pageWidth - marginX, y);
    y += 14;
  };

  const ensureSpace = (needed = 80) => {
    if (y + needed < pageHeight - marginY) return;
    doc.addPage();
    y = marginY;
  };

  // Header (Word-form style)
  doc.setFillColor(245, 247, 250);
  doc.rect(marginX, y - 18, maxWidth, 44, 'F');
  line('BDO Zimbabwe/Malawi — AML/CFT/CPF Client Onboarding Extract', 12, 16, 'bold');
  line(`Case ID: ${caseData.id}    Status: ${caseData.status}`, 10, 14);
  line(`Generated: ${new Date().toISOString().split('T')[0]}`, 9, 18);
  hr();

  // Section 1: Client identification
  line('1. Client Identification', 11, 16, 'bold');
  line(`Client name: ${safe(caseData.clientName)}`, 10, 14);
  line(`Client type: ${caseData.clientType}`, 10, 14);
  line(`Registration / ID number: ${safe(caseData.regNumber)}`, 10, 14);
  const addrLines = splitText(doc, `Registered address: ${safe(caseData.registeredAddress)}`, maxWidth);
  addrLines.forEach((l, idx) => line(idx === 0 ? l : `  ${l}`, 10, 14));
  const nobLines = splitText(doc, `Nature of business: ${safe(caseData.natureOfBusiness)}`, maxWidth);
  nobLines.forEach((l, idx) => line(idx === 0 ? l : `  ${l}`, 10, 14));
  const poeLines = splitText(doc, `Purpose of engagement: ${safe(caseData.purposeOfEngagement)}`, maxWidth);
  poeLines.forEach((l, idx) => line(idx === 0 ? l : `  ${l}`, 10, 14));
  line(`Office: ${caseData.office}`, 10, 18);
  hr();

  // Contact info
  line('Contact Information', 11, 16, 'bold');
  line(`Address: ${safe(caseData.contactInfo.address)}`, 10, 14);
  line(`Official email: ${safe(caseData.contactInfo.email)}`, 10, 14);
  line(`Phone: ${safe(caseData.contactInfo.phone)}`, 10, 18);
  hr();

  // Directors (compact)
  ensureSpace(140);
  line('2. Directors / Trustees / Senior Management', 11, 16, 'bold');
  if (!caseData.directors.length) {
    line('No director records captured.', 10, 18);
  } else {
    caseData.directors.slice(0, 12).forEach((d, i) => {
      const row = `${i + 1}. ${safe(d.fullName)} — ${safe(d.position)} — ${safe(d.nationality)} — ${safe(d.idNumber)} — ${safe(d.countryOfResidence)}`;
      splitText(doc, row, maxWidth).forEach((l, idx) => line(idx === 0 ? l : `   ${l}`, 10, 14));
      ensureSpace(40);
    });
    if (caseData.directors.length > 12) line(`(Truncated: ${caseData.directors.length - 12} more director(s) not shown)`, 9, 16);
  }
  hr();

  // Beneficial owners (compact)
  ensureSpace(140);
  line('3. Beneficial Owners (≥25%)', 11, 16, 'bold');
  if (!caseData.beneficialOwners.length) {
    line('No beneficial owner records captured.', 10, 18);
  } else {
    caseData.beneficialOwners.slice(0, 12).forEach((b, i) => {
      const row = `${i + 1}. ${safe(b.fullName)} — ${b.ownershipPercentage}% — ${safe(b.basisOfControl)} — ${safe(b.country)} — ${safe(b.verificationSource)}`;
      splitText(doc, row, maxWidth).forEach((l, idx) => line(idx === 0 ? l : `   ${l}`, 10, 14));
      ensureSpace(40);
    });
    if (caseData.beneficialOwners.length > 12) line(`(Truncated: ${caseData.beneficialOwners.length - 12} more owner(s) not shown)`, 9, 16);
  }
  hr();

  // Risk matrix + Section 3 indicators
  ensureSpace(220);
  line('4. Risk Assessment (Section 3)', 11, 16, 'bold');
  line(`Risk ratings — Client: ${caseData.riskRatings.client}; Geography: ${caseData.riskRatings.geography}; Product/Service: ${caseData.riskRatings.productService}; Delivery: ${caseData.riskRatings.deliveryChannel}; Payment: ${caseData.riskRatings.paymentMode}`, 10, 14);
  line(`Overall inherent risk rating: ${caseData.overallRiskRating}`, 10, 16, 'bold');
  splitText(doc, `Rationale: ${safe(caseData.riskRationale)}`, maxWidth).forEach((l, idx) => line(idx === 0 ? l : `  ${l}`, 10, 14));
  y += 8;

  line('A. Client risk indicators', 10, 14, 'bold');
  line(`${checkbox(caseData.riskIndicators.client.pepOrAssociate)} Politically Exposed Person (PEP) or close associate`, 10, 14);
  line(`${checkbox(caseData.riskIndicators.client.npo)} Non Profit Organisation (NPO)`, 10, 14);
  line(`${checkbox(caseData.riskIndicators.client.complexOwnership)} Complex or opaque ownership/control structure`, 10, 14);
  line(`${checkbox(caseData.riskIndicators.client.cashIntensive)} Cash intensive activities/payment method`, 10, 14);
  line(`${checkbox(caseData.riskIndicators.client.intermediaries)} Use of intermediaries, agents or nominees`, 10, 14);
  line(`Other: ${safe(caseData.riskIndicators.client.other) || '—'}`, 10, 18);

  line('B. Geographic risk indicators', 10, 14, 'bold');
  line(`${checkbox(caseData.riskIndicators.geography.fatfGreyListed)} FATF grey listed jurisdiction`, 10, 14);
  line(`${checkbox(caseData.riskIndicators.geography.fatfBlackListed)} FATF black listed jurisdiction`, 10, 14);
  line(`${checkbox(caseData.riskIndicators.geography.sanctionsExposed)} Sanctions exposed country`, 10, 14);
  line(`${checkbox(caseData.riskIndicators.geography.highCorruptionOrConflict)} High corruption or conflict affected area`, 10, 14);
  line(`Other: ${safe(caseData.riskIndicators.geography.other) || '—'}`, 10, 18);

  line('C. Product / service risk indicators', 10, 14, 'bold');
  line(`${checkbox(caseData.riskIndicators.productService.trustOrCompanyFormation)} Trust or company formation / administration`, 10, 14);
  line(`${checkbox(caseData.riskIndicators.productService.manageClientFundsOrAssets)} Management of client funds or assets`, 10, 14);
  line(`${checkbox(caseData.riskIndicators.productService.crossBorderTransactions)} Cross border transactions`, 10, 14);
  line(`${checkbox(caseData.riskIndicators.productService.highValueOrComplexTransactions)} High value or complex transactions`, 10, 14);
  line(`Other: ${safe(caseData.riskIndicators.productService.other) || '—'}`, 10, 18);
  hr();

  // Signatures timeline
  ensureSpace(180);
  line('5. Approvals / Sign-offs', 11, 16, 'bold');
  line(`Preparer: ${caseData.signatures.preparer ? `${caseData.signatures.preparer.sign} (${caseData.signatures.preparer.date})` : '—'}`, 10, 14);
  line(`Compliance / MLRO: ${caseData.signatures.compliance ? `${caseData.signatures.compliance.sign} (${caseData.signatures.compliance.date})` : '—'}`, 10, 14);
  line(`Risk Partner: ${caseData.signatures.riskPartner ? `${caseData.signatures.riskPartner.sign} (${caseData.signatures.riskPartner.date})` : '—'}`, 10, 14);
  line(`Engagement Partner: ${caseData.signatures.engagementPartner ? `${caseData.signatures.engagementPartner.sign} (${caseData.signatures.engagementPartner.date})` : '—'}`, 10, 18);

  // Footer
  doc.setFontSize(8);
  doc.setTextColor(120);
  doc.text(
    'This PDF extract is generated from the BDO AML onboarding application and is intended for internal compliance recordkeeping.',
    marginX,
    pageHeight - 28,
    { maxWidth }
  );

  const fileName = `BDO_AML_Onboarding_${caseData.id}_${new Date().toISOString().split('T')[0]}.pdf`;
  doc.save(fileName);
}

