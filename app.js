/**
 * AJ Estates Lease Wizard
 * Texas Residential Lease Generator
 * Version 1.0
 */

// Wait for docx library to be loaded
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
        AlignmentType, BorderStyle, HeadingLevel, PageBreak, Footer, PageNumber,
        convertInchesToTwip, UnderlineType, ShadingType } = docx;

// Hardcoded landlord information (FIX 1)
const LANDLORD = {
    name: "AJ Estates LLC",
    address: "700 Central Expressway S, Suite 400, Allen, TX 75013",
    phone: "469-408-0447",
    email: "info@ajrealestategroup.com",
    emergencyPhone: "469-408-0447"
};

// ============================================================================
// INITIALIZATION
// ============================================================================

document.addEventListener('DOMContentLoaded', function() {
    initializeForm();
    loadDraft();
    setupEventListeners();
    updateCalculations();
    updateMaxOccupants();
    updateAddendumStatus();
});

function initializeForm() {
    // Set default dates
    const today = new Date();
    const leaseStart = document.getElementById('leaseStartDate');
    const leaseEnd = document.getElementById('leaseEndDate');

    if (leaseStart && !leaseStart.value) {
        // Default to first of next month
        const nextMonth = new Date(today.getFullYear(), today.getMonth() + 1, 1);
        leaseStart.value = formatDateForInput(nextMonth);
    }

    if (leaseEnd && !leaseEnd.value) {
        // Default to one year from start
        const endDate = new Date(today.getFullYear() + 1, today.getMonth() + 1, 0);
        leaseEnd.value = formatDateForInput(endDate);
    }
}

function setupEventListeners() {
    // Form submission
    const form = document.getElementById('leaseForm');
    form.addEventListener('submit', handleFormSubmit);

    // Pet toggle
    const petsAllowed = document.getElementById('petsAllowed');
    petsAllowed.addEventListener('change', togglePetFields);

    // HOA toggle
    const hasHOA = document.getElementById('hasHOA');
    hasHOA.addEventListener('change', toggleHOAFields);

    // Year built - lead paint check
    const yearBuilt = document.getElementById('yearBuilt');
    yearBuilt.addEventListener('input', checkLeadPaint);

    // Monthly rent - late fee calculation
    const monthlyRent = document.getElementById('monthlyRent');
    monthlyRent.addEventListener('input', updateCalculations);

    // Tenant names and additional occupants - auto-calculate max occupants
    const tenant1Name = document.getElementById('tenant1Name');
    const tenant2Name = document.getElementById('tenant2Name');
    const tenant3Name = document.getElementById('tenant3Name');
    const numAdditionalOccupants = document.getElementById('numAdditionalOccupants');

    tenant1Name.addEventListener('input', updateMaxOccupants);
    tenant2Name.addEventListener('input', updateMaxOccupants);
    tenant3Name.addEventListener('input', updateMaxOccupants);
    numAdditionalOccupants.addEventListener('input', updateMaxOccupants);

    // Gas appliances toggle
    const hasGasAppliances = document.getElementById('hasGasAppliances');
    hasGasAppliances.addEventListener('change', updateAddendumStatus);

    // Carpet toggle
    const hasCarpet = document.getElementById('hasCarpet');
    hasCarpet.addEventListener('change', updateAddendumStatus);

    // Preview button
    const previewBtn = document.getElementById('previewBtn');
    previewBtn.addEventListener('click', showPreview);

    // Save draft button
    const saveDraftBtn = document.getElementById('saveDraftBtn');
    saveDraftBtn.addEventListener('click', saveDraft);

    // Modal controls
    const closePreview = document.getElementById('closePreview');
    const closePreviewBtn = document.getElementById('closePreviewBtn');
    const generateFromPreview = document.getElementById('generateFromPreview');

    closePreview.addEventListener('click', hidePreview);
    closePreviewBtn.addEventListener('click', hidePreview);
    generateFromPreview.addEventListener('click', () => {
        hidePreview();
        generateLease();
    });

    // Close modal on outside click
    const modal = document.getElementById('previewModal');
    modal.addEventListener('click', (e) => {
        if (e.target === modal) hidePreview();
    });
}

// ============================================================================
// TOGGLE FUNCTIONS
// ============================================================================

function togglePetFields() {
    const petFields = document.getElementById('petFields');
    const petsAllowed = document.getElementById('petsAllowed');
    petFields.style.display = petsAllowed.checked ? 'block' : 'none';
    updateAddendumStatus();
}

function toggleHOAFields() {
    const hoaFields = document.getElementById('hoaFields');
    const hasHOA = document.getElementById('hasHOA');
    hoaFields.style.display = hasHOA.checked ? 'block' : 'none';
    updateAddendumStatus();
}

function checkLeadPaint() {
    const yearBuilt = document.getElementById('yearBuilt').value;
    const requiresLeadPaint = yearBuilt && parseInt(yearBuilt) < 1978;

    // Show/hide warning
    const warning = document.getElementById('leadPaintWarning');
    warning.style.display = requiresLeadPaint ? 'block' : 'none';

    updateAddendumStatus();
}

// ============================================================================
// CALCULATIONS
// ============================================================================

function updateCalculations() {
    const monthlyRent = parseFloat(document.getElementById('monthlyRent').value) || 0;
    const lateFee = monthlyRent * 0.12;
    document.getElementById('lateFeeDisplay').textContent = formatCurrency(lateFee);
}

function updateMaxOccupants() {
    // Count tenants with names filled in
    const tenantCount = [
        document.getElementById('tenant1Name').value,
        document.getElementById('tenant2Name').value,
        document.getElementById('tenant3Name').value
    ].filter(name => name.trim() !== '').length;

    // Get number of additional occupants
    const additionalCount = parseInt(document.getElementById('numAdditionalOccupants').value) || 0;

    // Calculate total
    const total = tenantCount + additionalCount;

    // Update display (minimum 1)
    document.getElementById('maxOccupantsDisplay').textContent = Math.max(1, total);
}

function updateAddendumStatus() {
    const yearBuilt = parseInt(document.getElementById('yearBuilt').value) || 2000;
    const petsAllowed = document.getElementById('petsAllowed').checked;
    const hasHOA = document.getElementById('hasHOA').checked;
    const hasGas = document.getElementById('hasGasAppliances').checked;
    const hasCarpet = document.getElementById('hasCarpet').checked;

    updateAddendumItem('addendum-leadPaint', yearBuilt < 1978);
    updateAddendumItem('addendum-pet', petsAllowed);
    updateAddendumItem('addendum-hoa', hasHOA);
    updateAddendumItem('addendum-gas', hasGas);
    updateAddendumItem('addendum-carpet', hasCarpet);
}

function updateAddendumItem(id, isActive) {
    const item = document.getElementById(id);
    if (item) {
        if (isActive) {
            item.classList.add('active');
            item.querySelector('.addendum-status').textContent = '☑';
        } else {
            item.classList.remove('active');
            item.querySelector('.addendum-status').textContent = '☐';
        }
    }
}

// ============================================================================
// FORM VALIDATION
// ============================================================================

function validateForm() {
    const form = document.getElementById('leaseForm');
    const requiredFields = form.querySelectorAll('[required]');
    let isValid = true;
    let firstInvalid = null;

    requiredFields.forEach(field => {
        if (!field.value.trim()) {
            isValid = false;
            field.classList.add('invalid');
            if (!firstInvalid) firstInvalid = field;
        } else {
            field.classList.remove('invalid');
        }
    });

    if (!isValid && firstInvalid) {
        firstInvalid.focus();
        showToast('Please fill in all required fields', 'error');
    }

    return isValid;
}

function handleFormSubmit(e) {
    e.preventDefault();
    if (validateForm()) {
        generateLease();
    }
}

// ============================================================================
// DRAFT SAVE/LOAD
// ============================================================================

function saveDraft() {
    const formData = collectFormData();
    localStorage.setItem('leaseWizardDraft', JSON.stringify(formData));
    showToast('Draft saved successfully!', 'success');
}

function loadDraft() {
    const draft = localStorage.getItem('leaseWizardDraft');
    if (draft) {
        try {
            const data = JSON.parse(draft);
            populateForm(data);
            showToast('Draft loaded', 'success');
        } catch (e) {
            console.error('Error loading draft:', e);
        }
    }
}

function populateForm(data) {
    Object.keys(data).forEach(key => {
        const field = document.getElementById(key);
        if (field) {
            if (field.type === 'checkbox') {
                field.checked = data[key];
            } else {
                field.value = data[key];
            }
        }
    });

    // Trigger updates
    togglePetFields();
    toggleHOAFields();
    checkLeadPaint();
    updateCalculations();
    updateMaxOccupants();
    updateAddendumStatus();
}

// ============================================================================
// PREVIEW
// ============================================================================

function showPreview() {
    if (!validateForm()) return;

    const data = collectFormData();
    const previewContent = document.getElementById('previewContent');

    previewContent.innerHTML = generatePreviewHTML(data);
    document.getElementById('previewModal').classList.add('active');
}

function hidePreview() {
    document.getElementById('previewModal').classList.remove('active');
}

function generatePreviewHTML(data) {
    const lateFee = (parseFloat(data.monthlyRent) || 0) * 0.12;
    const monthToMonthRent = (parseFloat(data.monthlyRent) || 0) * 1.10;
    const tenants = [data.tenant1Name, data.tenant2Name, data.tenant3Name].filter(Boolean).join(', ');
    const garageText = data.garageSpaces == 1 ? '1-Car' : `${data.garageSpaces}-Car`;

    return `
        <div class="preview-section">
            <h3>Property Information</h3>
            <div class="preview-grid">
                <div class="preview-item"><span class="preview-label">Address:</span><span class="preview-value">${data.propertyAddress}</span></div>
                <div class="preview-item"><span class="preview-label">City:</span><span class="preview-value">${data.city}, TX ${data.zipCode}</span></div>
                <div class="preview-item"><span class="preview-label">County:</span><span class="preview-value">${data.county}</span></div>
                <div class="preview-item"><span class="preview-label">Beds/Baths:</span><span class="preview-value">${data.bedrooms} Bedrooms | ${data.bathrooms} Bathrooms</span></div>
                <div class="preview-item"><span class="preview-label">Garage:</span><span class="preview-value">${garageText} Garage</span></div>
                <div class="preview-item"><span class="preview-label">Year Built:</span><span class="preview-value">${data.yearBuilt || 'Not Available'}</span></div>
            </div>
        </div>

        <div class="preview-section">
            <h3>Tenant Information</h3>
            <div class="preview-grid">
                <div class="preview-item"><span class="preview-label">Tenant(s):</span><span class="preview-value">${tenants}</span></div>
                <div class="preview-item"><span class="preview-label">Phone:</span><span class="preview-value">${data.tenant1Phone || 'N/A'}</span></div>
                <div class="preview-item"><span class="preview-label">Email:</span><span class="preview-value">${data.tenant1Email || 'N/A'}</span></div>
                <div class="preview-item"><span class="preview-label">Max Occupants:</span><span class="preview-value">${data.maxOccupants}</span></div>
            </div>
        </div>

        <div class="preview-section">
            <h3>Lease Terms</h3>
            <div class="preview-grid">
                <div class="preview-item"><span class="preview-label">Start Date:</span><span class="preview-value">${formatDate(data.leaseStartDate)}</span></div>
                <div class="preview-item"><span class="preview-label">End Date:</span><span class="preview-value">${formatDate(data.leaseEndDate)}</span></div>
                <div class="preview-item"><span class="preview-label">Monthly Rent:</span><span class="preview-value">${formatCurrency(data.monthlyRent)}</span></div>
                <div class="preview-item"><span class="preview-label">Security Deposit:</span><span class="preview-value">${formatCurrency(data.securityDeposit)}</span></div>
                <div class="preview-item"><span class="preview-label">Late Fee:</span><span class="preview-value">${formatCurrency(lateFee)}</span></div>
                <div class="preview-item"><span class="preview-label">Prorated Amount:</span><span class="preview-value">${formatCurrency(data.proratedFirstMonth || 0)}</span></div>
            </div>
        </div>

        ${data.petsAllowed ? `
        <div class="preview-section">
            <h3>Pet Information</h3>
            <div class="preview-grid">
                <div class="preview-item"><span class="preview-label">Pet Deposit:</span><span class="preview-value">${formatCurrency(data.petDeposit)}</span></div>
                <div class="preview-item"><span class="preview-label">Monthly Pet Rent:</span><span class="preview-value">${formatCurrency(data.monthlyPetRent)}</span></div>
                ${data.pet1Name ? `<div class="preview-item"><span class="preview-label">Pet 1:</span><span class="preview-value">${data.pet1Name} (${data.pet1Type} - ${data.pet1Breed}, ${data.pet1Weight} lbs)</span></div>` : ''}
                ${data.pet2Name ? `<div class="preview-item"><span class="preview-label">Pet 2:</span><span class="preview-value">${data.pet2Name} (${data.pet2Type} - ${data.pet2Breed}, ${data.pet2Weight} lbs)</span></div>` : ''}
            </div>
        </div>
        ` : ''}

        <div class="preview-section">
            <h3>Applicable Addendums</h3>
            <ul style="list-style: none; padding: 0;">
                <li>☑ Move-In/Move-Out Condition Report (Always)</li>
                ${parseInt(data.yearBuilt) < 1978 ? '<li>☑ Lead-Based Paint Disclosure</li>' : ''}
                ${data.petsAllowed ? '<li>☑ Pet Agreement</li>' : ''}
                ${data.hasHOA ? '<li>☑ HOA Rules Addendum</li>' : ''}
                ${data.hasGasAppliances ? '<li>☑ Gas Pilot Light Notice</li>' : ''}
                <li>☑ Excess Water Bill / Leak (Always)</li>
                <li>☑ Clogged Drain (Always)</li>
                ${data.hasCarpet ? '<li>☑ Carpet Cleaning Agreement</li>' : ''}
                <li>☑ Electric Breaker / Fuse (Always)</li>
                <li>☑ Broken Window (Always)</li>
                <li>☑ HVAC Air Filter Maintenance (Always)</li>
                <li>☑ Winterization / Freeze Protection (Always)</li>
            </ul>
        </div>

        <div class="preview-section">
            <h3>Flood Disclosure</h3>
            <div class="preview-grid">
                <div class="preview-item"><span class="preview-label">In Floodplain:</span><span class="preview-value">${data.inFloodplain ? 'Yes' : 'No'}</span></div>
                <div class="preview-item"><span class="preview-label">Flood Damage (5 yrs):</span><span class="preview-value">${data.hasFloodDamage ? 'Yes' : 'No'}</span></div>
            </div>
        </div>
    `;
}

// ============================================================================
// DATA COLLECTION
// ============================================================================

function collectFormData() {
    return {
        // Property
        propertyAddress: document.getElementById('propertyAddress').value,
        city: document.getElementById('city').value,
        county: document.getElementById('county').value,
        zipCode: document.getElementById('zipCode').value,
        bedrooms: document.getElementById('bedrooms').value,
        bathrooms: document.getElementById('bathrooms').value,
        garageSpaces: document.getElementById('garageSpaces').value,
        yearBuilt: document.getElementById('yearBuilt').value,

        // Tenants
        tenant1Name: document.getElementById('tenant1Name').value,
        tenant1Phone: document.getElementById('tenant1Phone').value,
        tenant1Email: document.getElementById('tenant1Email').value,
        tenant2Name: document.getElementById('tenant2Name').value,
        tenant3Name: document.getElementById('tenant3Name').value,
        numAdditionalOccupants: document.getElementById('numAdditionalOccupants').value,
        additionalOccupants: document.getElementById('additionalOccupants').value,
        maxOccupants: document.getElementById('maxOccupantsDisplay').textContent,

        // Lease Terms
        leaseStartDate: document.getElementById('leaseStartDate').value,
        leaseEndDate: document.getElementById('leaseEndDate').value,
        monthlyRent: document.getElementById('monthlyRent').value,
        securityDeposit: document.getElementById('securityDeposit').value,
        proratedFirstMonth: document.getElementById('proratedFirstMonth').value,
        renewalIncrease: document.getElementById('renewalIncrease').value,

        // Pets
        petsAllowed: document.getElementById('petsAllowed').checked,
        petDeposit: document.getElementById('petDeposit').value,
        monthlyPetRent: document.getElementById('monthlyPetRent').value,
        pet1Type: document.getElementById('pet1Type').value,
        pet1Breed: document.getElementById('pet1Breed').value,
        pet1Name: document.getElementById('pet1Name').value,
        pet1Weight: document.getElementById('pet1Weight').value,
        pet2Type: document.getElementById('pet2Type').value,
        pet2Breed: document.getElementById('pet2Breed').value,
        pet2Name: document.getElementById('pet2Name').value,
        pet2Weight: document.getElementById('pet2Weight').value,

        // Features
        hasHOA: document.getElementById('hasHOA').checked,
        hoaName: document.getElementById('hoaName').value,
        hoaContact: document.getElementById('hoaContact').value,
        hoaRules: document.getElementById('hoaRules').value,
        hasGasAppliances: document.getElementById('hasGasAppliances').checked,
        hasCarpet: document.getElementById('hasCarpet').checked,
        hasRefrigerator: document.getElementById('hasRefrigerator').checked,
        hasDishwasher: document.getElementById('hasDishwasher').checked,
        hasMicrowave: document.getElementById('hasMicrowave').checked,
        hasWasher: document.getElementById('hasWasher').checked,
        hasDryer: document.getElementById('hasDryer').checked,
        hasGarbageDisposal: document.getElementById('hasGarbageDisposal').checked,
        hvacFilterSize: document.getElementById('hvacFilterSize').value,

        // Flood Disclosure (FIX 4)
        inFloodplain: document.getElementById('inFloodplain').checked,
        hasFloodDamage: document.getElementById('hasFloodDamage').checked
    };
}

// ============================================================================
// DOCUMENT GENERATION
// ============================================================================

async function generateLease() {
    if (!validateForm()) return;

    const generateBtn = document.getElementById('generateBtn');
    generateBtn.classList.add('loading');
    generateBtn.disabled = true;

    try {
        const data = collectFormData();
        console.log('Generating lease document...');

        const doc = createLeaseDocument(data);
        console.log('Document created, converting to blob...');

        const blob = await Packer.toBlob(doc);
        console.log('Blob created, size:', blob.size);

        // Create filename
        const tenantLastName = data.tenant1Name.split(' ').pop() || 'Tenant';
        const addressShort = data.propertyAddress.replace(/[^a-zA-Z0-9]/g, '_').substring(0, 30);
        const dateStr = new Date().toISOString().split('T')[0];
        const filename = `Lease_${addressShort}_${tenantLastName}_${dateStr}.docx`;

        // Download lease
        console.log('Downloading lease:', filename);
        downloadBlob(blob, filename);

        showToast('Lease generated successfully!', 'success');

    } catch (error) {
        console.error('Error generating lease:', error);
        showToast('Error generating lease: ' + error.message, 'error');
    } finally {
        generateBtn.classList.remove('loading');
        generateBtn.disabled = false;
    }
}

function createLeaseDocument(data) {
    const lateFee = (parseFloat(data.monthlyRent) || 0) * 0.12;
    const tenantNames = [data.tenant1Name, data.tenant2Name, data.tenant3Name].filter(Boolean).join(', ');
    const fullAddress = `${data.propertyAddress}, ${data.city}, TX ${data.zipCode}`;
    const requiresLeadPaint = parseInt(data.yearBuilt) < 1978;

    // Build appliances list
    const appliances = [];
    if (data.hasRefrigerator) appliances.push('Refrigerator');
    if (data.hasDishwasher) appliances.push('Dishwasher');
    if (data.hasMicrowave) appliances.push('Microwave');
    if (data.hasWasher) appliances.push('Washer');
    if (data.hasDryer) appliances.push('Dryer');
    if (data.hasGarbageDisposal) appliances.push('Garbage Disposal');

    const sections = [];

    // Title (V2 FIX 5 - Improved document presentation)
    sections.push(
        new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 100 },
            children: [
                new TextRun({ text: "AJ ESTATES LLC", bold: true, size: 32 })
            ]
        }),
        new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 300 },
            children: [
                new TextRun({ text: "TEXAS RESIDENTIAL LEASE AGREEMENT", bold: true, size: 28 })
            ]
        }),
        new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 200 },
            border: {
                bottom: { style: BorderStyle.SINGLE, size: 6, color: "CCCCCC" }
            },
            children: []
        }),
        createParagraph(''),
        createParagraph(`This Residential Lease Agreement ("Lease") is entered into on ${formatDate(new Date().toISOString().split('T')[0])}, by and between:`),
        createParagraph('')
    );

    // Parties (FIX 1 & 2 - Hardcoded landlord info with emergency contact)
    sections.push(
        createHeading('PARTIES'),
        createParagraph(`LANDLORD: ${LANDLORD.name}`),
        createParagraph(`Address: ${LANDLORD.address}`),
        createParagraph(`Phone: ${LANDLORD.phone}`),
        createParagraph(`Email: ${LANDLORD.email}`),
        createBoldParagraph(`Emergency Contact Phone: ${LANDLORD.emergencyPhone}`),
        createParagraph('(For conditions materially affecting physical health or safety of an ordinary tenant - Texas Property Code § 92.020)'),
        createParagraph(''),
        createParagraph(`TENANT(S): ${tenantNames}`),
        createParagraph(`Phone: ${data.tenant1Phone || ''}`),
        createParagraph(`Email: ${data.tenant1Email || ''}`),
        createParagraph('')
    );

    // Section 1: Property (V2 FIX 6 - Improved property description format)
    const garageText = data.garageSpaces == 1 ? '1-Car Garage' : `${data.garageSpaces}-Car Garage`;
    const yearBuiltText = data.yearBuilt ? data.yearBuilt : 'Not Available';
    sections.push(
        createHeading('1. PROPERTY'),
        createParagraph(`Landlord agrees to lease to Tenant and Tenant agrees to lease from Landlord the property located at:`),
        createParagraph(''),
        createBoldParagraph(`${fullAddress}`),
        createParagraph(`County: ${data.county}`),
        createParagraph(`Property Description: ${data.bedrooms} Bedrooms | ${data.bathrooms} Bathrooms | ${garageText}`),
        createParagraph(`Year Built: ${yearBuiltText}`),
        createParagraph('')
    );

    // Section 2: Term
    sections.push(
        createHeading('2. LEASE TERM'),
        createParagraph(`The lease term shall begin on ${formatDate(data.leaseStartDate)} and end on ${formatDate(data.leaseEndDate)}, unless terminated earlier in accordance with this Lease.`),
        createParagraph('')
    );

    // Section 3: Rent (FIX 6 - Added NSF fee)
    sections.push(
        createHeading('3. RENT'),
        createParagraph(`Monthly Rent: ${formatCurrency(data.monthlyRent)}`),
        createParagraph(`Due Date: Rent is due on the 1st day of each month.`),
        createParagraph(`Grace Period: Rent received after the 3rd day of the month will be considered late.`),
        createParagraph(`Late Fee: ${formatCurrency(lateFee)} (12% of monthly rent) will be charged for late payment.`),
        createParagraph(`NSF/Returned Check Fee: $30.00 (per Texas Business & Commerce Code § 3.506)`),
        data.proratedFirstMonth ? createParagraph(`Prorated First Month Rent: ${formatCurrency(data.proratedFirstMonth)}`) : createParagraph(''),
        createParagraph(''),
        createParagraph('Payment Methods: Rent shall be paid via:'),
        createParagraph('  • Online payment portal'),
        createParagraph('  • Certified check or money order'),
        createParagraph('  • ACH bank transfer'),
        createParagraph('')
    );

    // Section 4: Security Deposit
    sections.push(
        createHeading('4. SECURITY DEPOSIT'),
        createParagraph(`Security Deposit Amount: ${formatCurrency(data.securityDeposit)}`),
        createParagraph(''),
        createParagraph('The security deposit shall be held by Landlord as security for the faithful performance of Tenant\'s obligations under this Lease. The deposit shall be returned within 30 days after Tenant vacates the premises, less any deductions for:'),
        createParagraph('  • Unpaid rent or fees'),
        createParagraph('  • Damage beyond normal wear and tear'),
        createParagraph('  • Cleaning costs if premises are not left in clean condition'),
        createParagraph('  • Other charges as permitted by Texas law'),
        createParagraph('')
    );

    // Section 5: Occupants
    sections.push(
        createHeading('5. OCCUPANTS'),
        createParagraph(`Maximum Occupants: ${data.maxOccupants}`),
        createParagraph(`Authorized Occupants: ${tenantNames}${data.additionalOccupants ? ', ' + data.additionalOccupants : ''}`),
        createParagraph(''),
        createParagraph('Only the individuals listed above are authorized to occupy the premises. Guests may stay for no more than 14 consecutive days or 30 days total in any 12-month period without prior written consent from Landlord.'),
        createParagraph('')
    );

    // Section 6: Utilities (V2 FIX 1 - Expanded list)
    sections.push(
        createHeading('6. UTILITIES AND SERVICES'),
        createParagraph(''),
        createBoldParagraph('TENANT is responsible for:'),
        createParagraph('  • Electricity'),
        createParagraph('  • Gas'),
        createParagraph('  • Water/Sewer'),
        createParagraph('  • Trash Collection'),
        createParagraph('  • Internet/Cable'),
        createParagraph('  • Lawn Care and Landscaping Maintenance'),
        createParagraph('  • Pool Maintenance (if applicable)'),
        createParagraph('  • Pest Control'),
        createParagraph('  • Security/Alarm Monitoring (if applicable)'),
        createParagraph(''),
        createBoldParagraph('LANDLORD is responsible for:'),
        createParagraph('  • HOA Fees (if applicable)'),
        createParagraph(''),
        createParagraph('Tenant must transfer utilities into Tenant\'s name within 3 days of lease commencement. Failure to maintain required utilities may be grounds for lease termination.'),
        createParagraph('')
    );

    // Section 7: Appliances
    sections.push(
        createHeading('7. APPLIANCES AND FIXTURES'),
        createParagraph('The following appliances and fixtures are included with the property:'),
        createParagraph(`  ${appliances.length > 0 ? appliances.join(', ') : 'None specified'}`),
        createParagraph(''),
        createParagraph('Landlord shall maintain these appliances in working order. Tenant shall promptly report any malfunction or damage.'),
        createParagraph('')
    );

    // Section 8: Maintenance
    sections.push(
        createHeading('8. MAINTENANCE AND REPAIRS'),
        createParagraph('Tenant agrees to:'),
        createParagraph('  • Keep the premises clean and sanitary'),
        createParagraph('  • Dispose of garbage properly'),
        createParagraph('  • Not damage or misuse the premises'),
        createParagraph('  • Promptly report any maintenance issues or needed repairs'),
        createParagraph('  • Replace HVAC filters every 90 days (or more frequently if needed)'),
        createParagraph('  • Maintain the lawn and landscaping (if applicable)'),
        createParagraph(''),
        createParagraph('Landlord shall be responsible for major repairs and maintenance, including HVAC, plumbing, electrical systems, and structural components.'),
        createParagraph('')
    );

    // Section 8A: Repair Remedy Notice (FIX 3 - CRITICAL)
    sections.push(
        createHeading('8A. IMPORTANT NOTICE OF TENANT\'S REPAIR REMEDIES'),
        createBoldUnderlinedParagraph('(Texas Property Code §§ 92.056 and 92.0561)'),
        createParagraph(''),
        createBoldParagraph('If Landlord fails to repair a condition that materially affects the physical health or safety of an ordinary tenant after receiving proper written notice, Tenant may be entitled to the following remedies under Texas Property Code:'),
        createParagraph(''),
        createBoldParagraph('(1) Terminate the lease;'),
        createParagraph(''),
        createBoldParagraph('(2) Have the condition repaired or remedied and deduct the cost from rent (not to exceed one month\'s rent or $500, whichever is greater);'),
        createParagraph(''),
        createBoldParagraph('(3) Obtain a court order directing the Landlord to make repairs, reduce rent, pay damages, and pay Tenant\'s court costs and attorney\'s fees.'),
        createParagraph(''),
        createBoldParagraph('These remedies are subject to compliance with statutory notice requirements and other conditions specified in the Texas Property Code.'),
        createParagraph('')
    );

    // Section 9: Alterations
    sections.push(
        createHeading('9. ALTERATIONS'),
        createParagraph('Tenant shall not make any alterations, additions, or improvements to the premises without prior written consent from Landlord. This includes but is not limited to:'),
        createParagraph('  • Painting'),
        createParagraph('  • Installing fixtures'),
        createParagraph('  • Modifying landscaping'),
        createParagraph('  • Adding or removing appliances'),
        createParagraph('')
    );

    // Section 10: Entry by Landlord
    sections.push(
        createHeading('10. ENTRY BY LANDLORD'),
        createParagraph('Landlord may enter the premises with at least 24 hours\' notice for the following purposes:'),
        createParagraph('  • Inspections'),
        createParagraph('  • Repairs and maintenance'),
        createParagraph('  • Showing the property to prospective tenants or buyers'),
        createParagraph('  • Emergency situations (no notice required)'),
        createParagraph('')
    );

    // Section 11: Insurance
    sections.push(
        createHeading('11. INSURANCE'),
        createParagraph('Tenant is strongly encouraged to obtain renter\'s insurance to cover personal belongings and liability. Landlord\'s insurance does not cover Tenant\'s personal property.'),
        createParagraph('')
    );

    // Section 12: Default and Remedies
    sections.push(
        createHeading('12. DEFAULT AND REMEDIES'),
        createParagraph('If Tenant fails to comply with any term of this Lease, Landlord may:'),
        createParagraph('  • Provide written notice of violation'),
        createParagraph('  • Pursue eviction proceedings as permitted by Texas law'),
        createParagraph('  • Recover damages, including unpaid rent and costs'),
        createParagraph(''),
        createParagraph('Tenant shall have the right to cure any default within the time period required by Texas law.'),
        createParagraph('')
    );

    // Section 13: Termination (V2 FIX 2 - Month-to-month clause with 10% increase)
    const monthToMonthRent = (parseFloat(data.monthlyRent) || 0) * 1.10;
    sections.push(
        createHeading('13. TERMINATION AND RENEWAL'),
        createParagraph('This Lease shall terminate on the end date specified above. Either party must provide at least 60 days\' written notice before the lease end date if they do not wish to renew.'),
        createParagraph(''),
        createBoldParagraph('Month-to-Month Conversion:'),
        createParagraph(`If Tenant remains in the property after the lease term expires without signing a new lease, the tenancy shall automatically convert to a month-to-month lease. Upon conversion to month-to-month, the monthly rent shall increase by 10% (to ${formatCurrency(monthToMonthRent)}).`),
        createParagraph(''),
        createParagraph('Either party may terminate a month-to-month tenancy by providing at least 30 days\' written notice.'),
        createParagraph(''),
        createBoldParagraph('Renewal Terms:'),
        createParagraph(`If both parties agree to renew for another fixed term, rent may be increased by up to 10% upon renewal.`),
        createParagraph('')
    );

    // Section 14: Move-Out Procedures
    sections.push(
        createHeading('14. MOVE-OUT PROCEDURES'),
        createParagraph('Upon termination of this Lease, Tenant shall:'),
        createParagraph('  • Remove all personal belongings'),
        createParagraph('  • Return all keys and access devices'),
        createParagraph('  • Leave premises in clean condition'),
        createParagraph('  • Provide forwarding address for deposit return'),
        createParagraph('  • Complete move-out inspection with Landlord'),
        createParagraph('')
    );

    // Section 15: Pets
    if (data.petsAllowed) {
        sections.push(
            createHeading('15. PETS'),
            createParagraph('Pets ARE permitted on the premises subject to the Pet Agreement Addendum.'),
            createParagraph(`Pet Deposit: ${formatCurrency(data.petDeposit)}`),
            createParagraph(`Monthly Pet Rent: ${formatCurrency(data.monthlyPetRent)}`),
            createParagraph('')
        );
    } else {
        sections.push(
            createHeading('15. PETS'),
            createParagraph('NO PETS are permitted on the premises without prior written consent from Landlord. Service animals are exempt from this restriction with proper documentation.'),
            createParagraph('')
        );
    }

    // Section 16: HOA (V2 FIX 3 - Tenant contacts HOA for rules)
    if (data.hasHOA) {
        sections.push(
            createHeading('16. HOMEOWNERS ASSOCIATION'),
            createParagraph('☑ APPLIES TO THIS LEASE'),
            createParagraph(''),
            createParagraph(`This property is subject to the rules and regulations of: ${data.hoaName || 'the Homeowners Association'}`),
            createParagraph(''),
            createParagraph('HOA Contact: Contact Landlord for details'),
            createParagraph(''),
            createParagraph('Tenant agrees to comply with all HOA rules and regulations. Tenant may contact the HOA directly to obtain a copy of the current rules and regulations. Tenant shall be responsible for any fines or penalties resulting from Tenant\'s violation of HOA rules.'),
            createParagraph('')
        );
    } else {
        sections.push(
            createHeading('16. HOMEOWNERS ASSOCIATION'),
            createParagraph('☐ DOES NOT APPLY'),
            createParagraph(''),
            createParagraph('This property is NOT subject to HOA rules and regulations.'),
            createParagraph('')
        );
    }

    // Section 17: Flood Disclosure (FIX 4 - CRITICAL)
    sections.push(
        createHeading('17. FLOOD DISCLOSURE'),
        createBoldParagraph('(Texas Property Code § 92.0135)'),
        createParagraph(''),
        createParagraph('100-Year Floodplain Status:'),
        createParagraph(`  ${data.inFloodplain ? '☑' : '☐'} Landlord IS aware that the property is located in a 100-year floodplain`),
        createParagraph(`  ${data.inFloodplain ? '☐' : '☑'} Landlord IS NOT aware that the property is located in a 100-year floodplain`),
        createParagraph(''),
        createParagraph('Prior Flooding:'),
        createParagraph(`  ${data.hasFloodDamage ? '☑' : '☐'} Landlord IS aware that flooding has damaged the property during the previous 5 years`),
        createParagraph(`  ${data.hasFloodDamage ? '☐' : '☑'} Landlord IS NOT aware that flooding has damaged the property during the previous 5 years`),
        createParagraph('')
    );

    // Section 18: Special Termination Rights (FIX 5)
    sections.push(
        createHeading('18. SPECIAL TERMINATION RIGHTS'),
        createParagraph('Texas law provides certain tenants with the right to terminate a lease early without penalty in the following circumstances:'),
        createParagraph(''),
        createBoldParagraph('Family Violence (§ 92.016):'),
        createParagraph('Tenants who are victims of family violence may terminate the lease by providing documentation as specified in the statute.'),
        createParagraph(''),
        createBoldParagraph('Sexual Assault or Stalking (§ 92.0161):'),
        createParagraph('Tenants who are victims of sexual assault or stalking may terminate the lease by providing documentation as specified in the statute.'),
        createParagraph(''),
        createBoldParagraph('Military Deployment (§ 92.017):'),
        createParagraph('Service members who receive military orders for permanent change of station or deployment may terminate the lease in accordance with federal and state law.'),
        createParagraph(''),
        createBoldParagraph('Death of Sole Tenant (§ 92.0162):'),
        createParagraph('If the sole tenant dies, the tenant\'s estate or representative may terminate the lease in accordance with the statute.'),
        createParagraph(''),
        createParagraph('Contact Landlord for specific procedures and documentation requirements.'),
        createParagraph('')
    );

    // Section 19: Additional Terms
    sections.push(
        createHeading('19. ADDITIONAL TERMS'),
        createParagraph('  • No smoking is permitted inside the premises'),
        createParagraph('  • No illegal activities on the premises'),
        createParagraph('  • Tenant shall not disturb neighbors or engage in nuisance behavior'),
        createParagraph('  • Tenant shall comply with all applicable laws and ordinances'),
        createParagraph('')
    );

    // Section 20: Notices
    sections.push(
        createHeading('20. NOTICES'),
        createParagraph('All notices required under this Lease shall be in writing and delivered to the addresses listed above via:'),
        createParagraph('  • Personal delivery'),
        createParagraph('  • Certified mail, return receipt requested'),
        createParagraph('  • Email (with confirmation of receipt)'),
        createParagraph('')
    );

    // Section 21: Governing Law
    sections.push(
        createHeading('21. GOVERNING LAW'),
        createParagraph('This Lease shall be governed by the laws of the State of Texas. Any disputes arising under this Lease shall be resolved in the courts of ' + data.county + ' County, Texas.'),
        createParagraph('')
    );

    // Section 22: Severability
    sections.push(
        createHeading('22. SEVERABILITY'),
        createParagraph('If any provision of this Lease is found to be invalid or unenforceable, the remaining provisions shall continue in full force and effect.'),
        createParagraph('')
    );

    // Section 23: Entire Agreement
    sections.push(
        createHeading('23. ENTIRE AGREEMENT'),
        createParagraph('This Lease, together with all addendums, constitutes the entire agreement between the parties. No verbal agreements shall be binding.'),
        createParagraph('')
    );

    // Section 24: Addendums (FIX 9 - Updated list with all 12 addendums)
    sections.push(
        createHeading('24. ADDENDUMS'),
        createParagraph('The following addendums are incorporated into this Lease:'),
        createParagraph('  ☑ Move-In/Move-Out Condition Report (Always)'),
        createParagraph(`  ${requiresLeadPaint ? '☑' : '☐'} Lead-Based Paint Disclosure (If built before 1978)`),
        createParagraph(`  ${data.petsAllowed ? '☑' : '☐'} Pet Agreement (If pets allowed)`),
        createParagraph(`  ${data.hasHOA ? '☑' : '☐'} HOA Rules Addendum (If property has HOA)`),
        createParagraph(`  ${data.hasGasAppliances ? '☑' : '☐'} Gas Pilot Light Notice (If gas appliances)`),
        createParagraph('  ☑ Excess Water Bill / Leak (Always)'),
        createParagraph('  ☑ Clogged Drain (Always)'),
        createParagraph(`  ${data.hasCarpet ? '☑' : '☐'} Carpet Cleaning Agreement (If has carpet)`),
        createParagraph('  ☑ Electric Breaker / Fuse (Always)'),
        createParagraph('  ☑ Broken Window (Always)'),
        createParagraph('  ☑ HVAC Air Filter Maintenance (Always)'),
        createParagraph('  ☑ Winterization / Freeze Protection (Always)'),
        createParagraph('')
    );

    // Signatures
    sections.push(
        createHeading('25. SIGNATURES'),
        createParagraph('By signing below, the parties agree to all terms and conditions of this Lease.'),
        createParagraph(''),
        createParagraph(''),
        createParagraph('LANDLORD:'),
        createParagraph(''),
        createParagraph('_____________________________________________    Date: _______________'),
        createParagraph(`${LANDLORD.name}`),
        createParagraph(''),
        createParagraph(''),
        createParagraph('TENANT(S):'),
        createParagraph(''),
        createParagraph('_____________________________________________    Date: _______________'),
        createParagraph(data.tenant1Name),
        createParagraph('')
    );

    if (data.tenant2Name) {
        sections.push(
            createParagraph('_____________________________________________    Date: _______________'),
            createParagraph(data.tenant2Name),
            createParagraph('')
        );
    }

    if (data.tenant3Name) {
        sections.push(
            createParagraph('_____________________________________________    Date: _______________'),
            createParagraph(data.tenant3Name),
            createParagraph('')
        );
    }

    // Add addendums - ALWAYS include all addendums for consistent Dropbox Sign template
    sections.push(new Paragraph({ children: [new PageBreak()] }));

    // Lead Paint Addendum (always included, marked if applicable)
    sections.push(...createLeadPaintAddendum(data, requiresLeadPaint));

    // Pet Addendum (always included, marked if applicable)
    sections.push(...createPetAddendum(data, data.petsAllowed));

    // HOA Addendum (always included, marked if applicable)
    sections.push(...createHOAAddendum(data, data.hasHOA));

    // Gas Pilot Light Addendum (always included, marked if applicable)
    sections.push(...createGasAddendum(data, data.hasGasAppliances));

    // Water Leak Addendum (always applies)
    sections.push(...createWaterLeakAddendum(data));

    // Clogged Drain Addendum (always applies)
    sections.push(...createCloggedDrainAddendum(data));

    // Carpet Cleaning Addendum (always included, marked if applicable)
    sections.push(...createCarpetAddendum(data, data.hasCarpet));

    // Electric Breaker Addendum (always applies)
    sections.push(...createElectricBreakerAddendum(data));

    // Broken Window Addendum (always applies)
    sections.push(...createBrokenWindowAddendum(data));

    // HVAC Filter Addendum (always applies)
    sections.push(...createHVACFilterAddendum(data));

    // Winterization Addendum (always applies)
    sections.push(...createWinterizationAddendum(data));

    // Move-In Checklist
    sections.push(...createMoveInChecklist(data));

    // V2 FIX 4 - Add page footer with lease identification
    const leaseDate = formatDate(new Date().toISOString().split('T')[0]);
    const pageFooter = new Footer({
        children: [
            new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 100 },
                children: [
                    new TextRun({ text: "─".repeat(70), size: 16, color: "999999" })
                ]
            }),
            new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                    new TextRun({ text: `${LANDLORD.name} | ${fullAddress} | Lease Date: ${leaseDate} | Page `, size: 16 }),
                    new TextRun({ children: [PageNumber.CURRENT], size: 16 }),
                    new TextRun({ text: " of ", size: 16 }),
                    new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 16 })
                ]
            }),
            new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                    new TextRun({ text: "Tenant Initials: _______ | Landlord Initials: _______", size: 16 })
                ]
            })
        ]
    });

    return new Document({
        sections: [{
            properties: {},
            children: sections,
            footers: {
                default: pageFooter
            }
        }]
    });
}

// ============================================================================
// ADDENDUM GENERATORS
// ============================================================================

function createLeadPaintAddendum(data, applies = false) {
    const statusLine = applies
        ? '☑ APPLIES TO THIS LEASE (Property built before 1978)'
        : '☐ DOES NOT APPLY (Property built 1978 or later)';

    return [
        createHeading('ADDENDUM A: LEAD-BASED PAINT DISCLOSURE', true),
        createParagraph(''),
        createBoldParagraph(statusLine),
        createParagraph(''),
        createParagraph('Disclosure of Information on Lead-Based Paint and/or Lead-Based Paint Hazards'),
        createParagraph(''),
        createParagraph('Lead Warning Statement:'),
        createParagraph('Housing built before 1978 may contain lead-based paint. Lead from paint, paint chips, and dust can pose health hazards if not managed properly. Lead exposure is especially harmful to young children and pregnant women. Before renting pre-1978 housing, lessors must disclose the presence of known lead-based paint and/or lead-based paint hazards in the dwelling.'),
        createParagraph(''),
        createParagraph('Lessor\'s Disclosure:'),
        createParagraph('(a) Presence of lead-based paint and/or lead-based paint hazards (check one):'),
        createParagraph('    ☐ Known lead-based paint and/or lead-based paint hazards are present in the housing.'),
        createParagraph('    ☑ Lessor has no knowledge of lead-based paint and/or lead-based paint hazards in the housing.'),
        createParagraph(''),
        createParagraph('(b) Records and reports available to the lessor (check one):'),
        createParagraph('    ☐ Lessor has provided the lessee with all available records and reports pertaining to lead-based paint and/or lead-based paint hazards.'),
        createParagraph('    ☑ Lessor has no reports or records pertaining to lead-based paint and/or lead-based paint hazards in the housing.'),
        createParagraph(''),
        createParagraph('Lessee\'s Acknowledgment:'),
        createParagraph('(c) Lessee has received the pamphlet "Protect Your Family from Lead in Your Home."'),
        createParagraph('(d) Lessee has received a 10-day opportunity (or mutually agreed upon period) to conduct a risk assessment or inspection for the presence of lead-based paint and/or lead-based paint hazards.'),
        createParagraph(''),
        createParagraph('Agent\'s Acknowledgment:'),
        createParagraph('(e) Agent has informed the lessor of the lessor\'s obligations under 42 U.S.C. 4852(d) and is aware of his/her responsibility to ensure compliance.'),
        createParagraph(''),
        createParagraph(''),
        createParagraph('Lessor Signature: ___________________________________ Date: _______________'),
        createParagraph(''),
        createParagraph('Lessee Signature: ___________________________________ Date: _______________'),
        createParagraph(''),
        new Paragraph({ children: [new PageBreak()] })
    ];
}

function createPetAddendum(data, applies = false) {
    const statusLine = applies
        ? '☑ APPLIES TO THIS LEASE'
        : '☐ DOES NOT APPLY (No pets permitted without signed pet addendum)';

    const pets = [];
    if (applies && data.pet1Name) {
        pets.push(`${data.pet1Name} (${data.pet1Type} - ${data.pet1Breed}, ${data.pet1Weight} lbs)`);
    }
    if (applies && data.pet2Name) {
        pets.push(`${data.pet2Name} (${data.pet2Type} - ${data.pet2Breed}, ${data.pet2Weight} lbs)`);
    }

    const sections = [
        createHeading('ADDENDUM B: PET AGREEMENT', true),
        createParagraph(''),
        createBoldParagraph(statusLine),
        createParagraph(''),
        createParagraph('This Pet Agreement is attached to and made part of the Residential Lease Agreement.'),
        createParagraph('')
    ];

    if (applies && pets.length > 0) {
        sections.push(
            createParagraph('APPROVED PETS:'),
            ...pets.map(pet => createParagraph(`  • ${pet}`)),
            createParagraph(''),
            createParagraph(`Pet Deposit: ${formatCurrency(data.petDeposit)} (non-refundable)`),
            createParagraph(`Monthly Pet Rent: ${formatCurrency(data.monthlyPetRent)} per pet`),
            createParagraph('')
        );
    } else {
        sections.push(
            createParagraph('APPROVED PETS: None'),
            createParagraph(''),
            createParagraph('Pet Deposit: N/A'),
            createParagraph('Monthly Pet Rent: N/A'),
            createParagraph('')
        );
    }

    sections.push(
        createParagraph('TENANT AGREES TO:'),
        createParagraph('1. Keep pet(s) current on all vaccinations and provide proof upon request.'),
        createParagraph('2. Keep pet(s) on a leash when outside the dwelling unit.'),
        createParagraph('3. Immediately clean up after pet(s) and properly dispose of waste.'),
        createParagraph('4. Prevent pet(s) from causing damage to the premises or disturbance to neighbors.'),
        createParagraph('5. Not allow pet(s) to become a nuisance through excessive barking or aggressive behavior.'),
        createParagraph('6. Be liable for any damage caused by pet(s) to the premises or other property.'),
        createParagraph('7. Comply with all local animal ordinances.'),
        createParagraph('8. Not breed pets on the premises.'),
        createParagraph('9. Limit pets to those specifically listed in this agreement.'),
        createParagraph(''),
        createParagraph('LANDLORD RESERVES THE RIGHT TO:'),
        createParagraph('• Revoke pet privileges with 30 days\' notice if Tenant violates this agreement.'),
        createParagraph('• Require removal of any pet that becomes a nuisance or danger.'),
        createParagraph(''),
        createParagraph(''),
        createParagraph('Tenant Signature: ___________________________________ Date: _______________'),
        createParagraph(''),
        createParagraph('Landlord Signature: ___________________________________ Date: _______________'),
        createParagraph(''),
        new Paragraph({ children: [new PageBreak()] })
    );

    return sections;
}

function createHOAAddendum(data, applies = false) {
    const statusLine = applies
        ? '☑ APPLIES TO THIS LEASE'
        : '☐ DOES NOT APPLY (Property is not subject to HOA)';

    return [
        createHeading('ADDENDUM C: HOA RULES AND REGULATIONS', true),
        createParagraph(''),
        createBoldParagraph(statusLine),
        createParagraph(''),
        createParagraph('This property is subject to the rules and regulations of the Homeowners Association.'),
        createParagraph(''),
        createParagraph(`HOA Name: ${applies && data.hoaName ? data.hoaName : 'N/A'}`),
        createParagraph(`HOA Contact: ${applies && data.hoaContact ? data.hoaContact : 'N/A'}`),
        createParagraph(''),
        createParagraph('TENANT ACKNOWLEDGES AND AGREES:'),
        createParagraph(''),
        createParagraph('1. Tenant has received a copy of the HOA rules and regulations, or agrees to obtain them from the HOA directly.'),
        createParagraph(''),
        createParagraph('2. Tenant shall comply with all HOA rules, regulations, covenants, and restrictions.'),
        createParagraph(''),
        createParagraph('3. Tenant shall be responsible for any fines, penalties, or assessments imposed by the HOA as a result of Tenant\'s violation of HOA rules.'),
        createParagraph(''),
        createParagraph('4. Common HOA restrictions may include (but are not limited to):'),
        createParagraph('   • Parking restrictions'),
        createParagraph('   • Noise restrictions'),
        createParagraph('   • Landscaping requirements'),
        createParagraph('   • Exterior modification restrictions'),
        createParagraph('   • Pet restrictions'),
        createParagraph('   • Trash and recycling procedures'),
        createParagraph(''),
        applies && data.hoaRules ? createParagraph(`Additional HOA Notes: ${data.hoaRules}`) : createParagraph(''),
        createParagraph(''),
        createParagraph(''),
        createParagraph('Tenant Signature: ___________________________________ Date: _______________'),
        createParagraph(''),
        createParagraph('Landlord Signature: ___________________________________ Date: _______________'),
        createParagraph(''),
        new Paragraph({ children: [new PageBreak()] })
    ];
}

function createGasAddendum(data, applies = false) {
    // FIX 7 - Updated with $150 service fee language
    const statusLine = applies
        ? '☑ APPLIES TO THIS LEASE'
        : '☐ DOES NOT APPLY (No gas appliances)';

    return [
        createHeading('ADDENDUM D: GAS PILOT LIGHTING', true),
        createParagraph(''),
        createBoldParagraph(statusLine),
        createParagraph(''),
        createParagraph('Landlord is not responsible for lighting pilots on gas stoves, heaters, or water heaters. Directions to light all gas appliances are clearly written and mounted on each appliance.'),
        createParagraph(''),
        createParagraph('Tenant agrees to light and keep all pilots lit at all times.'),
        createParagraph(''),
        createBoldParagraph('If Tenant cannot light a pilot and requests Landlord or Landlord\'s service technician to do so, Tenant will be charged a $150 service fee if the technician is successful at lighting the pilot.'),
        createParagraph(''),
        createParagraph('Tenant agrees not to hold Landlord responsible for any damages or injuries due to lighting of any gas pilots.'),
        createParagraph(''),
        createParagraph(''),
        createBoldParagraph('GAS SAFETY GUIDELINES:'),
        createParagraph(''),
        createParagraph('1. CARBON MONOXIDE: Gas appliances can produce carbon monoxide, an odorless, colorless gas that can be deadly. Ensure all gas appliances are properly vented.'),
        createParagraph(''),
        createParagraph('2. GAS LEAKS: Natural gas has an added odorant that smells like rotten eggs. If you smell gas:'),
        createParagraph('   • Do NOT turn on/off any electrical switches or appliances'),
        createParagraph('   • Do NOT use any open flames'),
        createParagraph('   • Leave the premises immediately'),
        createParagraph('   • Call the gas company emergency line from outside'),
        createParagraph('   • Call 911 if necessary'),
        createParagraph(''),
        createParagraph('3. DETECTORS: Ensure carbon monoxide detectors are installed and functioning. Test monthly and replace batteries as needed.'),
        createParagraph(''),
        createParagraph('EMERGENCY CONTACTS:'),
        createParagraph('• Gas Company Emergency: 911 or local gas company'),
        createParagraph(`• Landlord: ${LANDLORD.emergencyPhone}`),
        createParagraph(''),
        createParagraph(''),
        createParagraph('Tenant Signature: ___________________________________ Date: _______________'),
        createParagraph(''),
        new Paragraph({ children: [new PageBreak()] })
    ];
}

function createCarpetAddendum(data, applies = false) {
    const statusLine = applies
        ? '☑ APPLIES TO THIS LEASE'
        : '☐ DOES NOT APPLY (No carpet in property)';

    return [
        createHeading('ADDENDUM E: CARPET CLEANING AGREEMENT', true),
        createParagraph(''),
        createBoldParagraph(statusLine),
        createParagraph(''),
        createParagraph('This property has carpeted flooring. Tenant acknowledges and agrees to the following:'),
        createParagraph(''),
        createParagraph('TENANT RESPONSIBILITIES:'),
        createParagraph(''),
        createParagraph('1. Tenant shall maintain carpets in clean condition throughout the tenancy.'),
        createParagraph(''),
        createParagraph('2. Tenant shall vacuum carpets regularly (at least weekly) to prevent dirt buildup.'),
        createParagraph(''),
        createParagraph('3. Tenant shall promptly clean any spills or stains to prevent permanent damage.'),
        createParagraph(''),
        createParagraph('4. Tenant shall not use bleach or harsh chemicals that may damage carpet fibers.'),
        createParagraph(''),
        createParagraph('MOVE-OUT REQUIREMENTS:'),
        createParagraph(''),
        createParagraph('1. Carpets must be professionally cleaned by a licensed carpet cleaning company prior to move-out.'),
        createParagraph(''),
        createParagraph('2. Tenant shall provide Landlord with a receipt from the professional cleaning service.'),
        createParagraph(''),
        createParagraph('3. If Tenant fails to professionally clean carpets, Landlord will arrange for cleaning and deduct the cost from the security deposit.'),
        createParagraph(''),
        createParagraph('4. Tenant shall be responsible for repair or replacement costs for any carpet damage beyond normal wear and tear, including but not limited to:'),
        createParagraph('   • Pet stains or odors'),
        createParagraph('   • Burns or tears'),
        createParagraph('   • Permanent stains'),
        createParagraph('   • Bleach damage'),
        createParagraph(''),
        createParagraph(''),
        createParagraph('Tenant Signature: ___________________________________ Date: _______________'),
        createParagraph(''),
        createParagraph('Landlord Signature: ___________________________________ Date: _______________'),
        createParagraph(''),
        new Paragraph({ children: [new PageBreak()] })
    ];
}

function createMoveInChecklist(data) {
    const rooms = ['Living Room', 'Kitchen', 'Master Bedroom', 'Bedroom 2', 'Bedroom 3', 'Master Bathroom', 'Bathroom 2', 'Garage', 'Exterior/Yard'];
    const conditions = ['Walls', 'Floors', 'Ceiling', 'Windows', 'Doors', 'Light Fixtures', 'Outlets', 'Other'];

    return [
        createHeading('MOVE-IN / MOVE-OUT CONDITION REPORT', true),
        createParagraph(''),
        createParagraph(`Property: ${data.propertyAddress}, ${data.city}, TX ${data.zipCode}`),
        createParagraph(`Tenant: ${data.tenant1Name}`),
        createParagraph(''),
        createParagraph('Instructions: Complete this form at move-in and move-out. Note the condition of each item using: E=Excellent, G=Good, F=Fair, P=Poor, N/A=Not Applicable'),
        createParagraph(''),
        ...rooms.flatMap(room => [
            createParagraph(`${room.toUpperCase()}:`),
            ...conditions.map(item => createParagraph(`  ${item}: Move-In: _____ | Move-Out: _____ | Notes: _________________`)),
            createParagraph('')
        ]),
        createParagraph(''),
        createParagraph('APPLIANCES:'),
        createParagraph('  Refrigerator: Move-In: _____ | Move-Out: _____ | Notes: _________________'),
        createParagraph('  Stove/Oven: Move-In: _____ | Move-Out: _____ | Notes: _________________'),
        createParagraph('  Dishwasher: Move-In: _____ | Move-Out: _____ | Notes: _________________'),
        createParagraph('  Microwave: Move-In: _____ | Move-Out: _____ | Notes: _________________'),
        createParagraph('  Washer: Move-In: _____ | Move-Out: _____ | Notes: _________________'),
        createParagraph('  Dryer: Move-In: _____ | Move-Out: _____ | Notes: _________________'),
        createParagraph(''),
        createParagraph('KEYS/ACCESS DEVICES RECEIVED:'),
        createParagraph('  House Keys: _____ | Mailbox Keys: _____ | Garage Remotes: _____'),
        createParagraph('  Gate Code: _____ | Other: _____'),
        createParagraph(''),
        createParagraph(''),
        createParagraph('MOVE-IN:'),
        createParagraph(`Date: _______________ Landlord: _______________ Tenant: _______________`),
        createParagraph(''),
        createParagraph('MOVE-OUT:'),
        createParagraph(`Date: _______________ Landlord: _______________ Tenant: _______________`),
        createParagraph('')
    ];
}

// FIX 8: NEW ADDENDUMS

function createWaterLeakAddendum(data) {
    return [
        createHeading('ADDENDUM: EXCESS WATER BILL DUE TO LEAK', true),
        createParagraph(''),
        createParagraph('☑ APPLIES TO THIS LEASE'),
        createParagraph(''),
        createParagraph('Tenant is responsible for water usage and payments at all times.'),
        createParagraph(''),
        createBoldParagraph('Tenant must notify Landlord IMMEDIATELY in writing upon seeing or hearing any water leak.'),
        createParagraph(''),
        createParagraph('Tenant is responsible for excess water charges in all cases where leaks are not promptly reported to Landlord in writing.'),
        createParagraph(''),
        createParagraph('Landlord has shown Tenant how to turn off the main water supply valve to the property. In case of a broken or burst water supply line, Tenant agrees to turn off the main water supply immediately before calling Landlord or a plumber.'),
        createParagraph(''),
        createBoldParagraph('REPORT ANY LEAKS TO LANDLORD IMMEDIATELY. OTHERWISE, TENANT WILL BE CHARGED FOR EXCESS WATER USAGE.'),
        createParagraph(''),
        createParagraph(''),
        createParagraph('Tenant Initials: _______     Date: _______'),
        createParagraph(''),
        new Paragraph({ children: [new PageBreak()] })
    ];
}

function createCloggedDrainAddendum(data) {
    return [
        createHeading('ADDENDUM: CLOGGED DRAIN', true),
        createParagraph(''),
        createParagraph('☑ APPLIES TO THIS LEASE'),
        createParagraph(''),
        createParagraph('Tenant acknowledges that during the pre-move-in inspection, all drains were tested and toilets flushed, and all were functioning properly and free of clogs.'),
        createParagraph(''),
        createBoldParagraph('Tenant agrees to pay for clearing all clogged drains, toilets, sinks, and traps caused by Tenant\'s use.'),
        createParagraph(''),
        createParagraph('In the event of a clog, Tenant must first attempt to clear the clog using a plunger or similar tool. If unsuccessful, Tenant will call a professional drain cleaning company (not the Landlord) to attempt to clear the line.'),
        createParagraph(''),
        createParagraph('If the professional drain cleaner cannot clear the clog and determines it is caused by a condition in the property\'s plumbing system (such as a broken pipe) not caused by Tenant, Landlord will reimburse Tenant for the service call and arrange for repairs.'),
        createParagraph(''),
        createParagraph('99% of clogs are caused by items flushed or poured down drains. Think before you flush!'),
        createParagraph(''),
        createParagraph(''),
        createParagraph('Tenant Initials: _______     Date: _______'),
        createParagraph(''),
        new Paragraph({ children: [new PageBreak()] })
    ];
}

function createElectricBreakerAddendum(data) {
    return [
        createHeading('ADDENDUM: ELECTRIC BREAKER AND FUSE', true),
        createParagraph(''),
        createParagraph('☑ APPLIES TO THIS LEASE'),
        createParagraph(''),
        createParagraph('Tenant agrees to take responsibility for resetting tripped breakers or replacing blown fuses.'),
        createParagraph(''),
        createParagraph('Landlord has instructed Tenant at move-in on how to reset tripped breakers and replace fuses.'),
        createParagraph(''),
        createBoldParagraph('Tenant agrees to check and reset breakers or replace fuses BEFORE calling Landlord for a service call.'),
        createParagraph(''),
        createBoldParagraph('If Landlord or Landlord\'s representative is called to the property and determines that the problem was a tripped breaker or blown fuse, Tenant will be charged a $150 service fee.'),
        createParagraph(''),
        createParagraph(''),
        createParagraph('Tenant Initials: _______     Date: _______'),
        createParagraph(''),
        new Paragraph({ children: [new PageBreak()] })
    ];
}

function createBrokenWindowAddendum(data) {
    return [
        createHeading('ADDENDUM: BROKEN WINDOW', true),
        createParagraph(''),
        createParagraph('☑ APPLIES TO THIS LEASE'),
        createParagraph(''),
        createParagraph('Tenant takes responsibility for all windows in the property.'),
        createParagraph(''),
        createParagraph('If any window becomes broken, cracked, or damaged during Tenant\'s occupancy for any reason (including acts of nature, attempted burglary, or unknown causes), Tenant is responsible for having the window repaired at Tenant\'s expense.'),
        createParagraph(''),
        createBoldParagraph('Tenant must report any broken window to Landlord within 24 hours.'),
        createParagraph(''),
        createParagraph('Tenant must arrange for repair of broken windows within 7 days of the damage occurring. Failure to repair broken windows in a timely manner may be considered a lease violation.'),
        createParagraph(''),
        createParagraph('This provision does not apply to window damage caused by Landlord\'s negligence or defects in the window that existed prior to Tenant\'s occupancy.'),
        createParagraph(''),
        createParagraph(''),
        createParagraph('Tenant Initials: _______     Date: _______'),
        createParagraph(''),
        new Paragraph({ children: [new PageBreak()] })
    ];
}

function createHVACFilterAddendum(data) {
    return [
        createHeading('ADDENDUM: HVAC AIR FILTER MAINTENANCE', true),
        createParagraph(''),
        createParagraph('☑ APPLIES TO THIS LEASE'),
        createParagraph(''),
        createBoldParagraph('Tenant must replace the HVAC air filter every 90 days (or more frequently if needed).'),
        createParagraph(''),
        createParagraph('Tenant must send photo evidence of the new filter to Landlord each time the filter is replaced.'),
        createParagraph(''),
        createParagraph(`Filter size for this property: ${data.hvacFilterSize || '___________________'}`),
        createParagraph(''),
        createBoldParagraph('Failure to replace air filters may result in HVAC damage. If HVAC repairs are needed due to clogged or dirty filters, Tenant will be responsible for the cost of repairs.'),
        createParagraph(''),
        createParagraph(''),
        createParagraph('Tenant Initials: _______     Date: _______'),
        createParagraph(''),
        new Paragraph({ children: [new PageBreak()] })
    ];
}

function createWinterizationAddendum(data) {
    return [
        createHeading('ADDENDUM: WINTERIZATION / FREEZE PROTECTION', true),
        createParagraph(''),
        createParagraph('☑ APPLIES TO THIS LEASE'),
        createParagraph(''),
        createBoldParagraph('During winter or any time freezing temperatures are forecast:'),
        createParagraph(''),
        createParagraph('• Tenant must cover all exterior faucets with outdoor foam faucet covers'),
        createParagraph('• Tenant must disconnect and drain all garden hoses from exterior faucets'),
        createParagraph('• Tenant must keep thermostat set to at least 55°F to prevent interior pipe freezing'),
        createParagraph('• If leaving property vacant for more than 48 hours during winter, Tenant should turn off main water valve'),
        createParagraph(''),
        createBoldParagraph('Landlord is NOT responsible for frozen/burst pipes or water damage if Tenant fails to take reasonable steps to winterize the property.'),
        createParagraph(''),
        createParagraph(''),
        createParagraph('Tenant Initials: _______     Date: _______'),
        createParagraph('')
    ];
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

function createHeading(text, isTitle = false) {
    // V2 FIX 5 - Improved section headers with better spacing
    return new Paragraph({
        children: [
            new TextRun({
                text: text,
                bold: true,
                size: isTitle ? 32 : 26  // Slightly larger section headers (13pt)
            })
        ],
        heading: isTitle ? HeadingLevel.TITLE : HeadingLevel.HEADING_1,
        alignment: isTitle ? AlignmentType.CENTER : AlignmentType.LEFT,
        spacing: { before: isTitle ? 0 : 300, after: 150 }  // Add space before sections
    });
}

function createParagraph(text) {
    return new Paragraph({
        children: [
            new TextRun({
                text: text,
                size: 22
            })
        ],
        spacing: { after: 100 }
    });
}

function createBoldParagraph(text) {
    return new Paragraph({
        children: [
            new TextRun({
                text: text,
                bold: true,
                size: 22
            })
        ],
        spacing: { after: 100 }
    });
}

function createBoldUnderlinedParagraph(text) {
    return new Paragraph({
        children: [
            new TextRun({
                text: text,
                bold: true,
                underline: { type: UnderlineType.SINGLE },
                size: 22
            })
        ],
        spacing: { after: 100 }
    });
}

// V2 FIX 5 - Shaded box for important notices
function createShadedNotice(text, isBold = true) {
    return new Paragraph({
        shading: { fill: "F5F5F5", type: ShadingType.CLEAR },
        border: {
            top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
            bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
            left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
            right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }
        },
        children: [
            new TextRun({
                text: text,
                bold: isBold,
                size: 22
            })
        ],
        spacing: { before: 100, after: 100 }
    });
}

function formatCurrency(value) {
    const num = parseFloat(value) || 0;
    return new Intl.NumberFormat('en-US', {
        style: 'currency',
        currency: 'USD'
    }).format(num);
}

function formatDate(dateStr) {
    if (!dateStr) return 'N/A';
    const date = new Date(dateStr + 'T00:00:00');
    return date.toLocaleDateString('en-US', {
        year: 'numeric',
        month: 'long',
        day: 'numeric'
    });
}

function formatDateForInput(date) {
    return date.toISOString().split('T')[0];
}

function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

function showToast(message, type = 'success') {
    // Remove existing toast
    const existingToast = document.querySelector('.toast');
    if (existingToast) {
        existingToast.remove();
    }

    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.textContent = message;
    document.body.appendChild(toast);

    setTimeout(() => {
        toast.remove();
    }, 3000);
}
