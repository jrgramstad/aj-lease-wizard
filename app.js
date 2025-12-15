/**
 * AJ Estates Lease Wizard
 * Texas Residential Lease Generator
 * Version 1.0
 */

// Wait for docx library to be loaded
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
        AlignmentType, BorderStyle, HeadingLevel, PageBreak,
        convertInchesToTwip, UnderlineType } = docx;

// ============================================================================
// INITIALIZATION
// ============================================================================

document.addEventListener('DOMContentLoaded', function() {
    initializeForm();
    loadDraft();
    setupEventListeners();
    updateCalculations();
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
    const warning = document.getElementById('leadPaintWarning');
    warning.style.display = yearBuilt && parseInt(yearBuilt) < 1978 ? 'block' : 'none';
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
    const tenants = [data.tenant1Name, data.tenant2Name, data.tenant3Name].filter(Boolean).join(', ');

    return `
        <div class="preview-section">
            <h3>Property Information</h3>
            <div class="preview-grid">
                <div class="preview-item"><span class="preview-label">Address:</span><span class="preview-value">${data.propertyAddress}</span></div>
                <div class="preview-item"><span class="preview-label">City:</span><span class="preview-value">${data.city}, TX ${data.zipCode}</span></div>
                <div class="preview-item"><span class="preview-label">County:</span><span class="preview-value">${data.county}</span></div>
                <div class="preview-item"><span class="preview-label">Beds/Baths:</span><span class="preview-value">${data.bedrooms} bed / ${data.bathrooms} bath</span></div>
                <div class="preview-item"><span class="preview-label">Garage:</span><span class="preview-value">${data.garageSpaces} spaces</span></div>
                <div class="preview-item"><span class="preview-label">Year Built:</span><span class="preview-value">${data.yearBuilt || 'N/A'}</span></div>
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
                ${parseInt(data.yearBuilt) < 1978 ? '<li>☑ Lead-Based Paint Disclosure</li>' : ''}
                ${data.petsAllowed ? '<li>☑ Pet Agreement</li>' : ''}
                ${data.hasHOA ? '<li>☑ HOA Rules Addendum</li>' : ''}
                ${data.hasGasAppliances ? '<li>☑ Gas Pilot Light Notice</li>' : ''}
                ${data.hasCarpet ? '<li>☑ Carpet Cleaning Agreement</li>' : ''}
                ${!(parseInt(data.yearBuilt) < 1978 || data.petsAllowed || data.hasHOA || data.hasGasAppliances || data.hasCarpet) ? '<li>No special addendums required</li>' : ''}
            </ul>
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
        additionalOccupants: document.getElementById('additionalOccupants').value,
        maxOccupants: document.getElementById('maxOccupants').value,

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

        // Landlord
        landlordName: document.getElementById('landlordName').value,
        landlordAddress: document.getElementById('landlordAddress').value,
        landlordPhone: document.getElementById('landlordPhone').value,
        landlordEmail: document.getElementById('landlordEmail').value
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
        const doc = createLeaseDocument(data);

        const blob = await Packer.toBlob(doc);

        // Create filename
        const tenantLastName = data.tenant1Name.split(' ').pop() || 'Tenant';
        const addressShort = data.propertyAddress.replace(/[^a-zA-Z0-9]/g, '_').substring(0, 30);
        const dateStr = new Date().toISOString().split('T')[0];
        const filename = `Lease_${addressShort}_${tenantLastName}_${dateStr}.docx`;

        // Download
        downloadBlob(blob, filename);

        showToast('Lease generated successfully!', 'success');

    } catch (error) {
        console.error('Error generating lease:', error);
        showToast('Error generating lease. Please try again.', 'error');
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

    // Title
    sections.push(
        createHeading('TEXAS RESIDENTIAL LEASE AGREEMENT', true),
        createParagraph(''),
        createParagraph(`This Residential Lease Agreement ("Lease") is entered into on ${formatDate(new Date().toISOString().split('T')[0])}, by and between:`),
        createParagraph('')
    );

    // Parties
    sections.push(
        createHeading('PARTIES'),
        createParagraph(`LANDLORD: ${data.landlordName || 'AJ Estates LLC'}`),
        createParagraph(`Address: ${data.landlordAddress || ''}`),
        createParagraph(`Phone: ${data.landlordPhone || ''}`),
        createParagraph(`Email: ${data.landlordEmail || ''}`),
        createParagraph(''),
        createParagraph(`TENANT(S): ${tenantNames}`),
        createParagraph(`Phone: ${data.tenant1Phone || ''}`),
        createParagraph(`Email: ${data.tenant1Email || ''}`),
        createParagraph('')
    );

    // Section 1: Property
    sections.push(
        createHeading('1. PROPERTY'),
        createParagraph(`Landlord agrees to lease to Tenant and Tenant agrees to lease from Landlord the property located at:`),
        createParagraph(`${fullAddress}`),
        createParagraph(`County: ${data.county}`),
        createParagraph(`Property Description: ${data.bedrooms} bedroom(s), ${data.bathrooms} bathroom(s), ${data.garageSpaces} garage space(s)`),
        createParagraph(`Year Built: ${data.yearBuilt || 'N/A'}`),
        createParagraph('')
    );

    // Section 2: Term
    sections.push(
        createHeading('2. LEASE TERM'),
        createParagraph(`The lease term shall begin on ${formatDate(data.leaseStartDate)} and end on ${formatDate(data.leaseEndDate)}, unless terminated earlier in accordance with this Lease.`),
        createParagraph('')
    );

    // Section 3: Rent
    sections.push(
        createHeading('3. RENT'),
        createParagraph(`Monthly Rent: ${formatCurrency(data.monthlyRent)}`),
        createParagraph(`Due Date: Rent is due on the 1st day of each month.`),
        createParagraph(`Grace Period: Rent received after the 3rd day of the month will be considered late.`),
        createParagraph(`Late Fee: ${formatCurrency(lateFee)} (12% of monthly rent) will be charged for late payment.`),
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

    // Section 6: Utilities
    sections.push(
        createHeading('6. UTILITIES AND SERVICES'),
        createParagraph('Tenant shall be responsible for payment of the following utilities and services:'),
        createParagraph('  • Electricity'),
        createParagraph('  • Gas'),
        createParagraph('  • Water/Sewer'),
        createParagraph('  • Trash collection'),
        createParagraph('  • Internet/Cable'),
        createParagraph(''),
        createParagraph('Tenant must transfer utilities into Tenant\'s name within 3 days of lease commencement.'),
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
        createParagraph('  • Replace HVAC filters monthly'),
        createParagraph('  • Maintain the lawn and landscaping (if applicable)'),
        createParagraph(''),
        createParagraph('Landlord shall be responsible for major repairs and maintenance, including HVAC, plumbing, electrical systems, and structural components.'),
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

    // Section 13: Termination
    sections.push(
        createHeading('13. TERMINATION AND RENEWAL'),
        createParagraph('This Lease shall terminate on the end date specified above. Either party must provide at least 60 days\' written notice before the lease end date if they do not wish to renew.'),
        createParagraph(''),
        createParagraph(`Renewal Terms: If both parties agree to renew, rent may be increased by up to ${data.renewalIncrease || 5}% upon renewal.`),
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

    // Section 16: HOA
    if (data.hasHOA) {
        sections.push(
            createHeading('16. HOMEOWNERS ASSOCIATION'),
            createParagraph(`This property is subject to the rules and regulations of: ${data.hoaName || 'the Homeowners Association'}`),
            createParagraph(`HOA Contact: ${data.hoaContact || 'Contact Landlord for details'}`),
            createParagraph(''),
            createParagraph('Tenant agrees to comply with all HOA rules and regulations. Tenant shall be responsible for any fines or penalties resulting from Tenant\'s violation of HOA rules.'),
            createParagraph('')
        );
    } else {
        sections.push(
            createHeading('16. HOMEOWNERS ASSOCIATION'),
            createParagraph('This property is NOT subject to HOA rules and regulations.'),
            createParagraph('')
        );
    }

    // Section 17: Additional Terms
    sections.push(
        createHeading('17. ADDITIONAL TERMS'),
        createParagraph('  • No smoking is permitted inside the premises'),
        createParagraph('  • No illegal activities on the premises'),
        createParagraph('  • Tenant shall not disturb neighbors or engage in nuisance behavior'),
        createParagraph('  • Tenant shall comply with all applicable laws and ordinances'),
        createParagraph('')
    );

    // Section 18: Notices
    sections.push(
        createHeading('18. NOTICES'),
        createParagraph('All notices required under this Lease shall be in writing and delivered to the addresses listed above via:'),
        createParagraph('  • Personal delivery'),
        createParagraph('  • Certified mail, return receipt requested'),
        createParagraph('  • Email (with confirmation of receipt)'),
        createParagraph('')
    );

    // Section 19: Governing Law
    sections.push(
        createHeading('19. GOVERNING LAW'),
        createParagraph('This Lease shall be governed by the laws of the State of Texas. Any disputes arising under this Lease shall be resolved in the courts of ' + data.county + ' County, Texas.'),
        createParagraph('')
    );

    // Section 20: Severability
    sections.push(
        createHeading('20. SEVERABILITY'),
        createParagraph('If any provision of this Lease is found to be invalid or unenforceable, the remaining provisions shall continue in full force and effect.'),
        createParagraph('')
    );

    // Section 21: Entire Agreement
    sections.push(
        createHeading('21. ENTIRE AGREEMENT'),
        createParagraph('This Lease, together with all addendums, constitutes the entire agreement between the parties. No verbal agreements shall be binding.'),
        createParagraph('')
    );

    // Section 22: Addendums
    sections.push(
        createHeading('22. ADDENDUMS'),
        createParagraph('The following addendums are incorporated into this Lease:'),
        createParagraph(`  ${requiresLeadPaint ? '☑' : '☐'} Lead-Based Paint Disclosure (Required for properties built before 1978)`),
        createParagraph(`  ${data.petsAllowed ? '☑' : '☐'} Pet Agreement`),
        createParagraph(`  ${data.hasHOA ? '☑' : '☐'} HOA Rules Addendum`),
        createParagraph(`  ${data.hasGasAppliances ? '☑' : '☐'} Gas Pilot Light Notice`),
        createParagraph(`  ${data.hasCarpet ? '☑' : '☐'} Carpet Cleaning Agreement`),
        createParagraph('  ☑ Move-In/Move-Out Condition Report'),
        createParagraph('')
    );

    // Signatures
    sections.push(
        createHeading('23. SIGNATURES'),
        createParagraph('By signing below, the parties agree to all terms and conditions of this Lease.'),
        createParagraph(''),
        createParagraph(''),
        createParagraph('LANDLORD:'),
        createParagraph(''),
        createParagraph('_____________________________________________    Date: _______________'),
        createParagraph(`${data.landlordName || 'AJ Estates LLC'}`),
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

    // Add addendums
    sections.push(new Paragraph({ children: [new PageBreak()] }));

    // Lead Paint Addendum
    if (requiresLeadPaint) {
        sections.push(...createLeadPaintAddendum(data));
    }

    // Pet Addendum
    if (data.petsAllowed) {
        sections.push(...createPetAddendum(data));
    }

    // HOA Addendum
    if (data.hasHOA) {
        sections.push(...createHOAAddendum(data));
    }

    // Gas Pilot Light Addendum
    if (data.hasGasAppliances) {
        sections.push(...createGasAddendum(data));
    }

    // Carpet Cleaning Addendum
    if (data.hasCarpet) {
        sections.push(...createCarpetAddendum(data));
    }

    // Move-In Checklist
    sections.push(...createMoveInChecklist(data));

    return new Document({
        sections: [{
            properties: {},
            children: sections
        }]
    });
}

// ============================================================================
// ADDENDUM GENERATORS
// ============================================================================

function createLeadPaintAddendum(data) {
    return [
        createHeading('ADDENDUM A: LEAD-BASED PAINT DISCLOSURE', true),
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

function createPetAddendum(data) {
    const pets = [];
    if (data.pet1Name) {
        pets.push(`${data.pet1Name} (${data.pet1Type} - ${data.pet1Breed}, ${data.pet1Weight} lbs)`);
    }
    if (data.pet2Name) {
        pets.push(`${data.pet2Name} (${data.pet2Type} - ${data.pet2Breed}, ${data.pet2Weight} lbs)`);
    }

    return [
        createHeading('ADDENDUM B: PET AGREEMENT', true),
        createParagraph(''),
        createParagraph('This Pet Agreement is attached to and made part of the Residential Lease Agreement.'),
        createParagraph(''),
        createParagraph('APPROVED PETS:'),
        ...pets.map(pet => createParagraph(`  • ${pet}`)),
        createParagraph(''),
        createParagraph(`Pet Deposit: ${formatCurrency(data.petDeposit)} (non-refundable)`),
        createParagraph(`Monthly Pet Rent: ${formatCurrency(data.monthlyPetRent)} per pet`),
        createParagraph(''),
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
    ];
}

function createHOAAddendum(data) {
    return [
        createHeading('ADDENDUM C: HOA RULES AND REGULATIONS', true),
        createParagraph(''),
        createParagraph('This property is subject to the rules and regulations of the Homeowners Association.'),
        createParagraph(''),
        createParagraph(`HOA Name: ${data.hoaName || 'See Landlord for details'}`),
        createParagraph(`HOA Contact: ${data.hoaContact || 'Contact Landlord'}`),
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
        data.hoaRules ? createParagraph(`Additional HOA Notes: ${data.hoaRules}`) : createParagraph(''),
        createParagraph(''),
        createParagraph(''),
        createParagraph('Tenant Signature: ___________________________________ Date: _______________'),
        createParagraph(''),
        createParagraph('Landlord Signature: ___________________________________ Date: _______________'),
        createParagraph(''),
        new Paragraph({ children: [new PageBreak()] })
    ];
}

function createGasAddendum(data) {
    return [
        createHeading('ADDENDUM D: GAS PILOT LIGHT NOTICE', true),
        createParagraph(''),
        createParagraph('IMPORTANT SAFETY INFORMATION'),
        createParagraph(''),
        createParagraph('This property has gas-powered appliances. Tenant acknowledges receipt of this notice regarding the safe operation of gas appliances.'),
        createParagraph(''),
        createParagraph('GAS SAFETY GUIDELINES:'),
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
        createParagraph('3. PILOT LIGHTS: Some gas appliances have pilot lights that should remain lit. Know the location of pilot lights and how to safely relight them if needed.'),
        createParagraph(''),
        createParagraph('4. DETECTORS: Ensure carbon monoxide detectors are installed and functioning. Test monthly and replace batteries as needed.'),
        createParagraph(''),
        createParagraph('5. MAINTENANCE: Report any issues with gas appliances immediately to the Landlord.'),
        createParagraph(''),
        createParagraph('EMERGENCY CONTACTS:'),
        createParagraph('• Gas Company Emergency: 911 or local gas company'),
        createParagraph(`• Landlord: ${data.landlordPhone || 'See lease for contact'}`),
        createParagraph(''),
        createParagraph(''),
        createParagraph('Tenant Signature: ___________________________________ Date: _______________'),
        createParagraph(''),
        new Paragraph({ children: [new PageBreak()] })
    ];
}

function createCarpetAddendum(data) {
    return [
        createHeading('ADDENDUM E: CARPET CLEANING AGREEMENT', true),
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

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

function createHeading(text, isTitle = false) {
    return new Paragraph({
        children: [
            new TextRun({
                text: text,
                bold: true,
                size: isTitle ? 32 : 24
            })
        ],
        heading: isTitle ? HeadingLevel.TITLE : HeadingLevel.HEADING_1,
        alignment: isTitle ? AlignmentType.CENTER : AlignmentType.LEFT,
        spacing: { after: 200 }
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
