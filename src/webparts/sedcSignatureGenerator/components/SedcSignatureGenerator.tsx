import * as React from 'react';
import { useState, useEffect } from 'react';
import styles from './SedcSignatureGenerator.module.scss';
import type { ISedcSignatureGeneratorProps } from './ISedcSignatureGeneratorProps';
import { PrimaryButton, IconButton, /*Toggle, TextField,*/ Spinner, MessageBar, MessageBarType } from '@fluentui/react';
import { GraphService, UserProfile } from '../../../services/GraphService';
import { SignatureTemplates, SignatureData } from '../../../templates/SignatureTemplates';
import { officeAddresses, officeTypeLabels, getFinalAddress } from '../../../data/OfficeAddresses';

const SedcSignatureGenerator: React.FC<ISedcSignatureGeneratorProps> = (props) => {
  // User profile state
  const [userProfile, setUserProfile] = useState<UserProfile | null>(null);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string>('');
  
  // Editable user info fields
  const [displayName, setDisplayName] = useState<string>('');
  const [jobTitle, setJobTitle] = useState<string>('');
  const [department, setDepartment] = useState<string>('');
  const [mail, setMail] = useState<string>('');
  const [businessPhone, setBusinessPhone] = useState<string>('');
  
  // Edit mode state for each field
  const [editingField, setEditingField] = useState<string | null>(null);
  const [tempValue, setTempValue] = useState<string>('');
  
  // Validation errors
  const [fieldErrors, setFieldErrors] = useState<Record<string, string>>({});
  
  //Office Addresses
  const [officeType, setOfficeType] = useState<string>('');
  const [specificLocation, setSpecificLocation] = useState<string>('');

  // Additional information
  const [unit, setUnit] = useState<string>('');
  const [personalMobile, setPersonalMobile] = useState<string>('');
  const [officeLocation, setOfficeLocation] = useState<string>('');
  
  // Social media toggle and fields
  // const [includeSocials, setIncludeSocials] = useState<boolean>(false);
  // const [personalLinkedIn, setPersonalLinkedIn] = useState<string>('');
  // const [personalFacebook, setPersonalFacebook] = useState<string>('');
  // const [personalInstagram, setPersonalInstagram] = useState<string>('');
  // const [personalTwitter, setPersonalTwitter] = useState<string>('');
  // const [personalTikTok, setPersonalTikTok] = useState<string>('');
  
  // Social media validation
  const [socialErrors /*setSocialErrors*/] = useState<Record<string, string>>({});
  
  // Other state
  const [signatureHtml, setSignatureHtml] = useState<string>('');
  const [copySuccess, setCopySuccess] = useState<boolean>(false);

  // Load user profile
  const loadUserProfile = async (): Promise<void> => {
    try {
      setLoading(true);
      const graphClient = await props.context.msGraphClientFactory.getClient('3');
      const graphService = new GraphService(graphClient);
      const profile = await graphService.getUserProfile();
      
      setUserProfile(profile);
      
      // Set editable fields with AAD data
      setDisplayName(profile.displayName || '');
      setJobTitle(profile.jobTitle || '');
      setDepartment(profile.department || '');
      setMail(profile.mail || '');
      // Business phone: AAD first, then default to 082551555
      setBusinessPhone(profile.businessPhones?.[0] || '082551555');
      setOfficeLocation(profile.officeLocation || '');
      
      setError('');
    } catch (err) {
      setError('Failed to load user profile. Please ensure API permissions are granted.');
      console.error(err);
    } finally {
      setLoading(false);
    }
  };

  // Validate URL format
  // const isValidUrl = (url: string, platform: string): boolean => {
  //   if (!url) return true;
    
  //   try {
  //     const urlObj = new URL(url);
  //     const hostname = urlObj.hostname.toLowerCase();
      
  //     switch (platform) {
  //       case 'linkedin':
  //         return hostname.indexOf('linkedin.com') !== -1;
  //       case 'facebook':
  //         return hostname.indexOf('facebook.com') !== -1 || hostname.indexOf('fb.com') !== -1;
  //       case 'instagram':
  //         return hostname.indexOf('instagram.com') !== -1;
  //       case 'twitter':
  //         return hostname.indexOf('twitter.com') !== -1 || hostname.indexOf('x.com') !== -1;
  //       case 'tiktok':
  //         return hostname.indexOf('tiktok.com') !== -1;
  //       default:
  //         return true;
  //     }
  //   } catch {
  //     return false;
  //   }
  // };

  // Validate social media URLs
  // const validateSocialUrl = (platform: string, url: string): void => {
  //   if (url && !isValidUrl(url, platform)) {
  //     setSocialErrors(prev => ({
  //       ...prev,
  //       [platform]: `Please enter a valid ${platform} URL`
  //     }));
  //   } else {
  //     setSocialErrors(prev => {
  //       const newErrors = { ...prev };
  //       delete newErrors[platform];
  //       return newErrors;
  //     });
  //   }
  // };

  // Generate signature
  const generateSignature = (): void => {
    // Validate required fields
    const errors: Record<string, string> = {};
    if (!displayName.trim()) errors.displayName = 'Name is required';
    if (!department.trim()) errors.department = 'Department is required';
    if (!mail.trim()) errors.mail = 'Email is required';
    
    setFieldErrors(errors);
    
    // Don't generate if there are errors
    if (Object.keys(errors).length > 0) {
      return;
    }

    const data: SignatureData = {
      displayName,
      jobTitle,
      department,
      companyName: userProfile?.companyName || 'SEDC',
      mail,
      businessPhones: businessPhone ? [businessPhone] : [],
      mobilePhone: personalMobile || '',
      officeLocation: officeLocation || '',
      photo: userProfile?.photo,
      includePhoto: true,
      includeOffice: !!officeLocation,
      unit: unit || undefined,
      personalMobile: personalMobile || undefined,
      // personalLinkedIn: includeSocials ? personalLinkedIn || undefined : undefined,
      // personalFacebook: includeSocials ? personalFacebook || undefined : undefined,
      // personalInstagram: includeSocials ? personalInstagram || undefined : undefined,
      // personalTwitter: includeSocials ? personalTwitter || undefined : undefined,
      // personalTikTok: includeSocials ? personalTikTok || undefined : undefined,
    };

    const html = SignatureTemplates.generateTemplate4(data);
    setSignatureHtml(html);
  };

  // Update officeLocation whenever officeType or specificLocation changes
  useEffect(() => {
    const address = getFinalAddress(officeType, specificLocation);
    setOfficeLocation(address);
  }, [officeType, specificLocation]);

  // Load profile on mount
  useEffect(() => {
    loadUserProfile().catch(console.error);
  }, []);

  // Regenerate signature when any field changes
  useEffect(() => {
    if (userProfile) {
      generateSignature();
    }
  }, [displayName, jobTitle, department, mail, businessPhone, unit, personalMobile, 
      officeLocation/*, includeSocials, personalLinkedIn, personalFacebook, 
      personalInstagram, personalTwitter, personalTikTok*/]);

  // Start editing a field
  const startEdit = (fieldName: string, currentValue: string): void => {
    setEditingField(fieldName);
    setTempValue(currentValue);
    // Clear error for this field when editing
    setFieldErrors(prev => {
      const newErrors = { ...prev };
      delete newErrors[fieldName];
      return newErrors;
    });
  };

  // Check if field can be retrieved from AAD
  const canRetrieveFromAAD = (fieldName: string): boolean => {
    const aadFields = ['displayName', 'jobTitle', 'department', 'mail', 'businessPhone', 'officeLocation'];
    return aadFields.indexOf(fieldName) !== -1;
  };

  // Auto-retrieve from AAD
  const autoRetrieveFromAAD = (fieldName: string): void => {
    if (!userProfile) return;

    switch (fieldName) {
      case 'displayName':
        setTempValue(userProfile.displayName || '');
        break;
      case 'jobTitle':
        setTempValue(userProfile.jobTitle || '');
        break;
      case 'department':
        setTempValue(userProfile.department || '');
        break;
      case 'mail':
        setTempValue(userProfile.mail || '');
        break;
      case 'businessPhone':
        setTempValue(userProfile.businessPhones?.[0] || '082551555');
        break;
      case 'officeLocation':
        setTempValue(userProfile.officeLocation || '');
        break;
    }
  };

  // Check if field is required
  const isRequiredField = (fieldName: string): boolean => {
    return ['displayName', 'department', 'mail'].indexOf(fieldName) !== -1;
  };

  // Delete field value
  const deleteField = (fieldName: string): void => {
    if (isRequiredField(fieldName)) {
      alert(`${fieldName === 'displayName' ? 'Name' : fieldName === 'mail' ? 'Email' : 'Department'} is a required field and cannot be deleted.`);
      return;
    }

    // Clear the temp value in edit mode
    setTempValue('');
    
    // Also clear the actual field value
    switch (fieldName) {
      case 'jobTitle':
        setJobTitle('');
        break;
      case 'businessPhone':
        setBusinessPhone('082551555'); // Default value
        break;
      case 'unit':
        setUnit('');
        break;
      case 'personalMobile':
        setPersonalMobile('');
        break;
      case 'officeLocation':
        setOfficeLocation('');
        break;
    }
  };

  // Save edited field
  const saveEdit = (fieldName: string): void => {
    // Validate required fields
    if (isRequiredField(fieldName) && !tempValue.trim()) {
      alert(`${fieldName === 'displayName' ? 'Name' : fieldName === 'mail' ? 'Email' : 'Department'} is required and cannot be empty.`);
      return;
    }

    switch (fieldName) {
      case 'displayName':
        setDisplayName(tempValue);
        break;
      case 'jobTitle':
        setJobTitle(tempValue);
        break;
      case 'department':
        setDepartment(tempValue);
        break;
      case 'mail':
        setMail(tempValue);
        break;
      case 'businessPhone':
        setBusinessPhone(tempValue);
        break;
      case 'unit':
        setUnit(tempValue);
        break;
      case 'personalMobile':
        setPersonalMobile(tempValue);
        break;
      case 'officeLocation':
        setOfficeLocation(tempValue);
        break;
    }
    setEditingField(null);
    setTempValue('');
  };

  // Cancel editing
  const cancelEdit = (): void => {
    setEditingField(null);
    setTempValue('');
  };

  // Copy signature to clipboard
  const copySignature = async (): Promise<void> => {
    try {
      const tempDiv = document.createElement('div');
      tempDiv.innerHTML = signatureHtml;
      document.body.appendChild(tempDiv);

      const range = document.createRange();
      range.selectNode(tempDiv);
      const selection = window.getSelection();
      if (selection) {
        selection.removeAllRanges();
        selection.addRange(range);
      }

      document.execCommand('copy');

      document.body.removeChild(tempDiv);
      if (selection) {
        selection.removeAllRanges();
      }

      setCopySuccess(true);
      setTimeout(() => setCopySuccess(false), 3000);
    } catch (err) {
      console.error('Failed to copy:', err);
      alert('Failed to copy signature. Please try selecting and copying manually.');
    }
  };

  // Render editable field
  const renderEditableField = (label: string, fieldName: string, value: string, isRequired: boolean = false): JSX.Element => {
  const isEditing = editingField === fieldName;
  const hasError = fieldErrors[fieldName];
  const canRetrieve = canRetrieveFromAAD(fieldName);
  const canDelete = !isRequired;

  return (
    <div style={{ marginBottom: '15px' }}>
      <div style={{ display: 'flex', alignItems: 'center' }}>
        <div style={{ flex: 1, display: 'flex', alignItems: 'center' }}>
          <strong style={{ minWidth: '120px' }}>{label}{isRequired && <span style={{ color: 'red' }}>*</span>}:</strong>
          {isEditing ? (
            <input
              type="text"
              value={tempValue}
              onChange={(e) => setTempValue(e.target.value)}
              style={{ 
                flex: 1,
                padding: '5px',
                border: '1px solid #ccc',
                borderRadius: '2px',
                marginLeft: '10px'
              }}
              autoFocus
            />
          ) : (
            <span style={{ color: hasError ? 'red' : 'inherit', marginLeft: '10px' }}>
              {value || '(Not set)'}
            </span>
          )}
        </div>
        <div style={{ marginLeft: '10px', display: 'flex', gap: '2px', alignItems: 'center' }}>
          {isEditing ? (
            <>
              {canRetrieve && (
                <IconButton
                  iconProps={{ iconName: 'Refresh' }}
                  title="Refresh"
                  onClick={() => autoRetrieveFromAAD(fieldName)}
                  styles={{ root: { color: '#0078d4' } }}
                />
              )}
              {canDelete && (
                <IconButton
                  iconProps={{ iconName: 'Delete' }}
                  title="Delete"
                  onClick={() => deleteField(fieldName)}
                  styles={{ root: { color: '#d13438' } }}
                />
              )}
              <IconButton
                iconProps={{ iconName: 'Accept' }}
                title="Save"
                onClick={() => saveEdit(fieldName)}
                styles={{ root: { color: '#107c10' } }}
              />
              <IconButton
                iconProps={{ iconName: 'Cancel' }}
                title="Cancel"
                onClick={cancelEdit}
                styles={{ root: { color: '#d13438' } }}
              />
            </>
          ) : (
            <IconButton
              iconProps={{ iconName: 'Edit' }}
              title="Edit"
              onClick={() => startEdit(fieldName, value)}
            />
          )}
        </div>
      </div>
      {hasError && (
        <div style={{ color: 'red', fontSize: '12px', marginTop: '5px', marginLeft: '130px' }}>
          {hasError}
        </div>
      )}
    </div>
  );
};

  if (loading) {
    return (
      <div className={styles.sedcSignatureGenerator}>
        <Spinner label="Loading your profile..." />
      </div>
    );
  }

  if (error) {
    return (
      <div className={styles.sedcSignatureGenerator}>
        <MessageBar messageBarType={MessageBarType.error}>
          {error}
        </MessageBar>
        <PrimaryButton text="Retry" onClick={loadUserProfile} style={{ marginTop: '10px' }} />
      </div>
    );
  }

  return (
    <div className={styles.sedcSignatureGenerator}>
      <div className={styles.container}>
        <h2 className={styles.title}>SEDC Email Signature Generator</h2>
        
        {copySuccess && (
          <MessageBar messageBarType={MessageBarType.success} isMultiline={false}>
            Signature copied to clipboard! Now paste it into Outlook settings.
          </MessageBar>
        )}

        <div className={styles.grid}>
          {/* Left Panel - Controls */}
          <div className={styles.controls}>
            
          {/* User Information Section */}
          <h3>User Information</h3>
          <div className={styles.section}>
            {renderEditableField('Name', 'displayName', displayName, true)}
            {renderEditableField('Title', 'jobTitle', jobTitle)}
            {renderEditableField('Unit', 'unit', unit)}
            {renderEditableField('Department', 'department', department, true)}
            {renderEditableField('Email', 'mail', mail, true)}
            {renderEditableField('Business Phone', 'businessPhone', businessPhone)}
          </div>

          {/* Additional Information Section */}
          <h3>Additional Information</h3>
          <div className={styles.section}>
            {renderEditableField('Personal Mobile', 'personalMobile', personalMobile)}
            
            {/* Office Location Selector */}
            <div style={{ marginBottom: '15px' }}>
              <div style={{ marginBottom: '10px' }}>
                <label><strong>Office Type:</strong></label>
                <select 
                  value={officeType}
                  onChange={(e) => {
                    setOfficeType(e.target.value);
                    setSpecificLocation(''); // Reset specific location when type changes
                  }}
                  style={{ width: '100%', padding: '8px', marginTop: '5px', border: '1px solid #ccc', borderRadius: '2px' }}
                >
                  <option value="">-- Select Office Type --</option>
                  {Object.keys(officeAddresses).map((key) => (
                    <option key={key} value={key}>
                      {officeTypeLabels[key]}
                    </option>
                  ))}
                </select>
              </div>

              {/* Show second dropdown for RO */}
              {officeType === 'RO' && officeAddresses.RO.locations && (
                <div style={{ marginBottom: '10px' }}>
                  <label><strong>Select Regional Office:</strong></label>
                  <select
                    value={specificLocation}
                    onChange={(e) => setSpecificLocation(e.target.value)}
                    style={{ width: '100%', padding: '8px', marginTop: '5px', border: '1px solid #ccc', borderRadius: '2px' }}
                  >
                    <option value="">-- Select RO Location --</option>
                    {Object.keys(officeAddresses.RO.locations).map((key) => (
                      <option key={key} value={key}>
                        {key.charAt(0) + key.slice(1).toLowerCase()}
                      </option>
                    ))}
                  </select>
                </div>
              )}

              {/* Show second dropdown for PIBU */}
              {officeType === 'PIBU' && officeAddresses.PIBU.locations && (
                <div style={{ marginBottom: '10px' }}>
                  <label><strong>Select PIBU Location:</strong></label>
                  <select
                    value={specificLocation}
                    onChange={(e) => setSpecificLocation(e.target.value)}
                    style={{ width: '100%', padding: '8px', marginTop: '5px', border: '1px solid #ccc', borderRadius: '2px' }}
                  >
                    <option value="">-- Select PIBU Location --</option>
                    {Object.keys(officeAddresses.PIBU.locations).map((key) => (
                      <option key={key} value={key}>
                        {key.charAt(0) + key.slice(1).toLowerCase()}
                      </option>
                    ))}
                  </select>
                </div>
              )}

              {/* Preview selected location */}
              {officeLocation && (
                <div style={{ marginTop: '10px', padding: '10px', backgroundColor: '#f0f0f0', borderRadius: '4px' }}>
                  <strong>Selected Location:</strong>
                  <div style={{ fontSize: '9pt', marginTop: '5px', color: '#333' }}>
                    📍 {officeLocation}
                  </div>
                </div>
              )}
            </div>
          </div>

            {/* Personal Social Media Section
            <h3>Personal Social Media</h3>
            <div className={styles.section}>
              <Toggle
                label="Include Personal Socials"
                checked={includeSocials}
                onChange={(_, checked) => setIncludeSocials(checked || false)}
              />
              
              {includeSocials && (
                <div style={{ marginTop: '15px' }}>
                  <TextField
                    label="LinkedIn"
                    placeholder="https://linkedin.com/in/yourname"
                    value={personalLinkedIn}
                    onChange={(_, value) => {
                      setPersonalLinkedIn(value || '');
                      validateSocialUrl('linkedin', value || '');
                    }}
                    errorMessage={socialErrors.linkedin}
                  />
                  <TextField
                    label="Facebook"
                    placeholder="https://facebook.com/yourname"
                    value={personalFacebook}
                    onChange={(_, value) => {
                      setPersonalFacebook(value || '');
                      validateSocialUrl('facebook', value || '');
                    }}
                    errorMessage={socialErrors.facebook}
                  />
                  <TextField
                    label="Instagram"
                    placeholder="https://instagram.com/yourname"
                    value={personalInstagram}
                    onChange={(_, value) => {
                      setPersonalInstagram(value || '');
                      validateSocialUrl('instagram', value || '');
                    }}
                    errorMessage={socialErrors.instagram}
                  />
                  <TextField
                    label="Twitter (X)"
                    placeholder="https://x.com/yourname"
                    value={personalTwitter}
                    onChange={(_, value) => {
                      setPersonalTwitter(value || '');
                      validateSocialUrl('twitter', value || '');
                    }}
                    errorMessage={socialErrors.twitter}
                  />
                  <TextField
                    label="TikTok"
                    placeholder="https://tiktok.com/@yourname"
                    value={personalTikTok}
                    onChange={(_, value) => {
                      setPersonalTikTok(value || '');
                      validateSocialUrl('tiktok', value || '');
                    }}
                    errorMessage={socialErrors.tiktok}
                  />
                </div>
              )}
            </div> */}

            <PrimaryButton
              text="Copy Signature"
              onClick={copySignature}
              iconProps={{ iconName: 'Copy' }}
              style={{ marginTop: '20px' }}
              disabled={Object.keys(fieldErrors).length > 0 || Object.keys(socialErrors).length > 0}
            />
          </div>

          {/* Right Panel - Preview */}
          <div className={styles.preview}>
            <h3>Preview</h3>
            <div
              className={styles.previewBox}
              dangerouslySetInnerHTML={{ __html: signatureHtml }}
            />
            
            <div className={styles.instructions}>
              <h4>How to Install in Outlook:</h4>
              <ol>
                <li>Click &quot;Copy Signature&quot; button above</li>
                <li>Open Outlook → File → Options → Mail → Signatures</li>
                <li>Click &quot;New&quot; to create a new signature</li>
                <li>Paste (Ctrl+V) into the signature editor</li>
                <li>Click &quot;OK&quot; to save</li>
              </ol>
            </div>
          </div>
        </div>
      </div>
    </div>

  );
};

export default SedcSignatureGenerator;