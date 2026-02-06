import { useState } from 'react';
import './PMDocumentForm.css';

const today = () => new Date().toISOString().split('T')[0];

const emptyImpactRow = () => ({ application: '', components: '', remarks: '' });
const emptyRiskRow = () => ({ assumptions: '', risks: '', otherImpacts: '', remarks: '' });
const emptyDetailRow = () => ({ application: '', components: '', regressionNeeded: 'No', dbImpact: 'No', effort: '', remarks: '' });

function PMDocumentForm({ isOpen, onClose, onGenerate, initialData }) {
  const defaults = initialData || {};

  const [form, setForm] = useState({
    pmNumber: defaults.pmNumber || '',
    crNumber: defaults.crNumber || '',
    issueDescription: defaults.issueDescription || '',
    preparedByName: defaults.preparedByName || '',
    preparedByRole: defaults.preparedByRole || '',
    preparedByDate: defaults.preparedByDate || today(),
    reviewedByName: defaults.reviewedByName || '',
    reviewedByRole: defaults.reviewedByRole || '',
    reviewedByDate: defaults.reviewedByDate || today(),
    approvedByName: defaults.approvedByName || '',
    approvedByRole: defaults.approvedByRole || '',
    approvedByDate: defaults.approvedByDate || today(),
    systemImpacts: defaults.systemImpacts?.length > 0 ? defaults.systemImpacts : [emptyImpactRow()],
    risks: defaults.risks?.length > 0 ? defaults.risks : [emptyRiskRow()],
    analysisDetails: defaults.analysisDetails?.length > 0 ? defaults.analysisDetails : [emptyDetailRow()],
    versionNumber: defaults.versionNumber || 'V1.0.0',
    changesMade: defaults.changesMade || 'Initial baseline created',
    changedBy: defaults.changedBy || '',
    effectiveDate: defaults.effectiveDate || today(),
  });

  const [expandedSections, setExpandedSections] = useState({
    people: true, intro: true, impact: true, risk: true, details: false, changelog: false,
  });

  const toggleSection = (key) => setExpandedSections(prev => ({ ...prev, [key]: !prev[key] }));

  const updateField = (field, value) => setForm(prev => ({ ...prev, [field]: value }));

  const updateRow = (arrayField, index, field, value) => {
    setForm(prev => {
      const arr = [...prev[arrayField]];
      arr[index] = { ...arr[index], [field]: value };
      return { ...prev, [arrayField]: arr };
    });
  };

  const addRow = (arrayField, emptyFn) => {
    setForm(prev => ({ ...prev, [arrayField]: [...prev[arrayField], emptyFn()] }));
  };

  const removeRow = (arrayField, index) => {
    setForm(prev => {
      const arr = prev[arrayField].filter((_, i) => i !== index);
      return { ...prev, [arrayField]: arr.length > 0 ? arr : [arrayField === 'systemImpacts' ? emptyImpactRow() : arrayField === 'risks' ? emptyRiskRow() : emptyDetailRow()] };
    });
  };

  const handleGenerate = (mode) => {
    onGenerate(form, mode); // mode: 'download' | 'upload'
  };

  if (!isOpen) return null;

  return (
    <div className="pm-form-overlay" onClick={onClose}>
      <div className="pm-form-modal" onClick={(e) => e.stopPropagation()}>
        <div className="pm-form-header">
          <h3>PM Impact Analysis Document</h3>
          <button className="pm-form-close" onClick={onClose}>&times;</button>
        </div>

        <div className="pm-form-body">
          {/* People Section */}
          <div className="pm-section">
            <button className="pm-section-toggle" onClick={() => toggleSection('people')}>
              <span className="pm-section-arrow">{expandedSections.people ? '\u25BC' : '\u25B6'}</span>
              Prepared / Reviewed / Approved By
            </button>
            {expandedSections.people && (
              <div className="pm-section-content">
                <div className="pm-field-grid pm-grid-3">
                  <div className="pm-field-group">
                    <label>Prepared By</label>
                    <input placeholder="Name" value={form.preparedByName} onChange={e => updateField('preparedByName', e.target.value)} />
                    <input placeholder="Role" value={form.preparedByRole} onChange={e => updateField('preparedByRole', e.target.value)} />
                    <input type="date" value={form.preparedByDate} onChange={e => updateField('preparedByDate', e.target.value)} />
                  </div>
                  <div className="pm-field-group">
                    <label>Reviewed By</label>
                    <input placeholder="Name" value={form.reviewedByName} onChange={e => updateField('reviewedByName', e.target.value)} />
                    <input placeholder="Role" value={form.reviewedByRole} onChange={e => updateField('reviewedByRole', e.target.value)} />
                    <input type="date" value={form.reviewedByDate} onChange={e => updateField('reviewedByDate', e.target.value)} />
                  </div>
                  <div className="pm-field-group">
                    <label>Approved By</label>
                    <input placeholder="Name" value={form.approvedByName} onChange={e => updateField('approvedByName', e.target.value)} />
                    <input placeholder="Role" value={form.approvedByRole} onChange={e => updateField('approvedByRole', e.target.value)} />
                    <input type="date" value={form.approvedByDate} onChange={e => updateField('approvedByDate', e.target.value)} />
                  </div>
                </div>
              </div>
            )}
          </div>

          {/* Introduction */}
          <div className="pm-section">
            <button className="pm-section-toggle" onClick={() => toggleSection('intro')}>
              <span className="pm-section-arrow">{expandedSections.intro ? '\u25BC' : '\u25B6'}</span>
              1.0 Introduction
            </button>
            {expandedSections.intro && (
              <div className="pm-section-content">
                <div className="pm-field-grid pm-grid-2">
                  <div className="pm-field-group">
                    <label>PM Number</label>
                    <input value={form.pmNumber} onChange={e => updateField('pmNumber', e.target.value)} placeholder="e.g. 13366" />
                  </div>
                  <div className="pm-field-group">
                    <label>CR Number</label>
                    <input value={form.crNumber} onChange={e => updateField('crNumber', e.target.value)} placeholder="e.g. 19078" />
                  </div>
                </div>
                <label>Issue Description</label>
                <textarea rows={4} value={form.issueDescription} onChange={e => updateField('issueDescription', e.target.value)} placeholder="Describe the issue or change..." />
              </div>
            )}
          </div>

          {/* Expected System Impact */}
          <div className="pm-section">
            <button className="pm-section-toggle" onClick={() => toggleSection('impact')}>
              <span className="pm-section-arrow">{expandedSections.impact ? '\u25BC' : '\u25B6'}</span>
              2.0 Expected System Impact
            </button>
            {expandedSections.impact && (
              <div className="pm-section-content">
                {form.systemImpacts.map((row, i) => (
                  <div key={i} className="pm-row-group">
                    <div className="pm-row-header">
                      <span>Impact #{i + 1}</span>
                      {form.systemImpacts.length > 1 && (
                        <button className="pm-remove-btn" onClick={() => removeRow('systemImpacts', i)}>Remove</button>
                      )}
                    </div>
                    <input placeholder="Impacted Application(s)" value={row.application} onChange={e => updateRow('systemImpacts', i, 'application', e.target.value)} />
                    <textarea rows={2} placeholder="Impacted File(s) / Component(s) Name" value={row.components} onChange={e => updateRow('systemImpacts', i, 'components', e.target.value)} />
                    <textarea rows={2} placeholder="Remarks" value={row.remarks} onChange={e => updateRow('systemImpacts', i, 'remarks', e.target.value)} />
                  </div>
                ))}
                <button className="pm-add-btn" onClick={() => addRow('systemImpacts', emptyImpactRow)}>+ Add Impact Row</button>
              </div>
            )}
          </div>

          {/* Assumptions and Risk */}
          <div className="pm-section">
            <button className="pm-section-toggle" onClick={() => toggleSection('risk')}>
              <span className="pm-section-arrow">{expandedSections.risk ? '\u25BC' : '\u25B6'}</span>
              3.0 Assumptions and Risk
            </button>
            {expandedSections.risk && (
              <div className="pm-section-content">
                {form.risks.map((row, i) => (
                  <div key={i} className="pm-row-group">
                    <div className="pm-row-header">
                      <span>Entry #{i + 1}</span>
                      {form.risks.length > 1 && (
                        <button className="pm-remove-btn" onClick={() => removeRow('risks', i)}>Remove</button>
                      )}
                    </div>
                    <textarea rows={2} placeholder="Assumptions" value={row.assumptions} onChange={e => updateRow('risks', i, 'assumptions', e.target.value)} />
                    <textarea rows={2} placeholder="Risk(s)" value={row.risks} onChange={e => updateRow('risks', i, 'risks', e.target.value)} />
                    <textarea rows={2} placeholder="Other Impact(s)" value={row.otherImpacts} onChange={e => updateRow('risks', i, 'otherImpacts', e.target.value)} />
                    <textarea rows={2} placeholder="Remarks" value={row.remarks} onChange={e => updateRow('risks', i, 'remarks', e.target.value)} />
                  </div>
                ))}
                <button className="pm-add-btn" onClick={() => addRow('risks', emptyRiskRow)}>+ Add Risk Row</button>
              </div>
            )}
          </div>

          {/* Impact Analysis Details */}
          <div className="pm-section">
            <button className="pm-section-toggle" onClick={() => toggleSection('details')}>
              <span className="pm-section-arrow">{expandedSections.details ? '\u25BC' : '\u25B6'}</span>
              4.0 Impact Analysis Details
            </button>
            {expandedSections.details && (
              <div className="pm-section-content">
                {form.analysisDetails.map((row, i) => (
                  <div key={i} className="pm-row-group">
                    <div className="pm-row-header">
                      <span>Detail #{i + 1}</span>
                      {form.analysisDetails.length > 1 && (
                        <button className="pm-remove-btn" onClick={() => removeRow('analysisDetails', i)}>Remove</button>
                      )}
                    </div>
                    <input placeholder="Impacted Application(s)" value={row.application} onChange={e => updateRow('analysisDetails', i, 'application', e.target.value)} />
                    <textarea rows={2} placeholder="Impacted File(s) / Component(s) Name" value={row.components} onChange={e => updateRow('analysisDetails', i, 'components', e.target.value)} />
                    <div className="pm-field-grid pm-grid-3">
                      <div className="pm-field-group">
                        <label>Regression Testing</label>
                        <select value={row.regressionNeeded} onChange={e => updateRow('analysisDetails', i, 'regressionNeeded', e.target.value)}>
                          <option value="Yes">Yes</option>
                          <option value="No">No</option>
                        </select>
                      </div>
                      <div className="pm-field-group">
                        <label>Database Impact</label>
                        <select value={row.dbImpact} onChange={e => updateRow('analysisDetails', i, 'dbImpact', e.target.value)}>
                          <option value="Yes">Yes</option>
                          <option value="No">No</option>
                        </select>
                      </div>
                      <div className="pm-field-group">
                        <label>Effort (hrs)</label>
                        <input type="number" placeholder="0" value={row.effort} onChange={e => updateRow('analysisDetails', i, 'effort', e.target.value)} />
                      </div>
                    </div>
                    <textarea rows={2} placeholder="Remarks" value={row.remarks} onChange={e => updateRow('analysisDetails', i, 'remarks', e.target.value)} />
                  </div>
                ))}
                <button className="pm-add-btn" onClick={() => addRow('analysisDetails', emptyDetailRow)}>+ Add Detail Row</button>
              </div>
            )}
          </div>

          {/* Change Log */}
          <div className="pm-section">
            <button className="pm-section-toggle" onClick={() => toggleSection('changelog')}>
              <span className="pm-section-arrow">{expandedSections.changelog ? '\u25BC' : '\u25B6'}</span>
              5.0 Change Log
            </button>
            {expandedSections.changelog && (
              <div className="pm-section-content">
                <div className="pm-field-grid pm-grid-2">
                  <div className="pm-field-group">
                    <label>Version Number</label>
                    <input value={form.versionNumber} onChange={e => updateField('versionNumber', e.target.value)} />
                  </div>
                  <div className="pm-field-group">
                    <label>Changed By</label>
                    <input value={form.changedBy} onChange={e => updateField('changedBy', e.target.value)} />
                  </div>
                </div>
                <label>Changes Made</label>
                <input value={form.changesMade} onChange={e => updateField('changesMade', e.target.value)} />
                <label>Effective Date</label>
                <input type="date" value={form.effectiveDate} onChange={e => updateField('effectiveDate', e.target.value)} />
              </div>
            )}
          </div>
        </div>

        <div className="pm-form-footer">
          <button className="pm-cancel-btn" onClick={onClose}>Cancel</button>
          <button className="pm-download-btn" onClick={() => handleGenerate('download')}>Download .docx</button>
          <button className="pm-upload-btn" onClick={() => handleGenerate('upload')}>Download &amp; Upload</button>
        </div>
      </div>
    </div>
  );
}

export default PMDocumentForm;
