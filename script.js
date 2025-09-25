// LDCC1 Inspection Report JavaScript Functions

function navigate(elementId) {
    // Toggle visibility of the associated div
    const targetDiv = document.getElementById('d' + elementId);
    if (targetDiv) {
        if (targetDiv.style.display === 'none' || targetDiv.style.display === '') {
            targetDiv.style.display = 'block';
            updatePreview(elementId);
        } else {
            targetDiv.style.display = 'none';
            clearPreview();
        }
    }
    
    // Update checkbox state
    const checkbox = document.getElementById(elementId);
    if (checkbox) {
        checkbox.checked = targetDiv && targetDiv.style.display === 'block';
    }
}

function updatePreview(elementId) {
    const targetDiv = document.getElementById('d' + elementId);
    const preview = document.getElementById('preview');
    
    if (targetDiv && preview) {
        // Extract the content from the target div
        const problemDescription = targetDiv.querySelector('.problem-description');
        const location = targetDiv.querySelector('.location');
        
        if (problemDescription && location) {
            let content = '<div class="preview-content">';
            content += '<h5>Location:</h5>';
            content += location.innerHTML;
            content += '<h5>Description:</h5>';
            content += problemDescription.innerHTML;
            content += '</div>';
            
            preview.innerHTML = content;
        }
    }
}

function clearPreview() {
    const preview = document.getElementById('preview');
    if (preview) {
        preview.innerHTML = 'Select a problem element in tree';
    }
}

// Expand/collapse all functionality
function expandAll() {
    const allCheckboxes = document.querySelectorAll('input[type="checkbox"]');
    const allDivs = document.querySelectorAll('div[id^="d"]');
    
    allCheckboxes.forEach(checkbox => checkbox.checked = true);
    allDivs.forEach(div => div.style.display = 'block');
}

function collapseAll() {
    const allCheckboxes = document.querySelectorAll('input[type="checkbox"]');
    const allDivs = document.querySelectorAll('div[id^="d"]');
    
    allCheckboxes.forEach(checkbox => checkbox.checked = false);
    allDivs.forEach(div => div.style.display = 'none');
    clearPreview();
}

// Initialize page
document.addEventListener('DOMContentLoaded', function() {
    console.log('LDCC1 Inspection Report initialized');
    
    // Add expand/collapse controls
    const inspectionTree = document.getElementById('inspection-tree');
    if (inspectionTree) {
        const controls = document.createElement('div');
        controls.style.cssText = 'margin-bottom: 10px; padding: 10px; background-color: #f0f0f0; border-radius: 3px;';
        controls.innerHTML = `
            <button onclick="expandAll()" style="margin-right: 10px; padding: 5px 10px;">Expand All</button>
            <button onclick="collapseAll()" style="padding: 5px 10px;">Collapse All</button>
        `;
        inspectionTree.insertBefore(controls, inspectionTree.firstChild);
    }
    
    // Set up click handlers for checkboxes
    const checkboxes = document.querySelectorAll('input[type="checkbox"][onclick]');
    checkboxes.forEach(checkbox => {
        // Remove inline onclick and add proper event listener
        const onclickAttr = checkbox.getAttribute('onclick');
        if (onclickAttr) {
            checkbox.removeAttribute('onclick');
            checkbox.addEventListener('change', function() {
                const match = onclickAttr.match(/navigate\((\d+)\)/);
                if (match) {
                    navigate(match[1]);
                }
            });
        }
    });
});

// Utility function to get summary statistics
function getSummaryStats() {
    const warnings = document.querySelectorAll('span[style*="background:#f2f794"]').length;
    const weakWarnings = document.querySelectorAll('span:contains("WEAK WARNING")').length;
    const errors = document.querySelectorAll('span:contains("ERROR")').length;
    
    return {
        warnings: warnings,
        weakWarnings: weakWarnings,
        errors: errors,
        total: warnings + weakWarnings + errors
    };
}

// Export functions for external use
window.LDCC1InspectionReport = {
    navigate: navigate,
    expandAll: expandAll,
    collapseAll: collapseAll,
    getSummaryStats: getSummaryStats
};