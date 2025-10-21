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
  
  // UI state for guide and modals
  const [showQuickStart, setShowQuickStart] = useState<boolean>(true);
  const [showHelpModal, setShowHelpModal] = useState<boolean>(false);
  const [showSuccessCard, setShowSuccessCard] = useState<boolean>(false);
  const [copyButtonText, setCopyButtonText] = useState<string>('Copy Signature');

  //Office Addresses
  const [officeType, setOfficeType] = useState<string>('');
  const [specificLocation, setSpecificLocation] = useState<string>('');

  const [isAddressCustomized, setIsAddressCustomized] = useState<boolean>(false);
  const [originalOfficeLocation, setOriginalOfficeLocation] = useState<string>('');

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

  // Check if all required fields are filled
  const areRequiredFieldsFilled = (): boolean => {
    return !!(displayName && department && mail);
  };

  // Get list of missing required fields
  const getMissingFields = (): string[] => {
    const missing: string[] = [];
    if (!displayName) missing.push('Name');
    if (!department) missing.push('Department');
    if (!mail) missing.push('Email');
    return missing;
  };
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
    setOriginalOfficeLocation(address); // Store original
    setIsAddressCustomized(false); // Reset customization flag
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

  // Office address specific functions
  const startEditOfficeAddress = (): void => {
    setEditingField('officeAddress');
    setTempValue(officeLocation);
  };

  const saveOfficeAddress = (): void => {
    setOfficeLocation(tempValue);
    setIsAddressCustomized(tempValue !== originalOfficeLocation);
    setEditingField(null);
    setTempValue('');
  };

  const refreshOfficeAddress = (): void => {
    setTempValue(originalOfficeLocation);
    setOfficeLocation(originalOfficeLocation);
    setIsAddressCustomized(false);
  };

  const deleteOfficeAddress = (): void => {
    setOfficeLocation('');
    setOfficeType('');
    setSpecificLocation('');
    setOriginalOfficeLocation('');
    setIsAddressCustomized(false);
    setEditingField(null);
    setTempValue('');
  };

  // Cancel editing
  const cancelEdit = (): void => {
    setEditingField(null);
    setTempValue('');
  };

  // Copy signature to clipboard
  const copySignature = (): void => {
  if (signatureHtml) {
    // Create a temporary div to hold the HTML
    const tempDiv = document.createElement('div');
    tempDiv.innerHTML = signatureHtml;
    document.body.appendChild(tempDiv);

    // Select and copy
    const range = document.createRange();
    range.selectNodeContents(tempDiv);
    const selection = window.getSelection();
    if (selection) {
      selection.removeAllRanges();
      selection.addRange(range);
      document.execCommand('copy');
      selection.removeAllRanges();
    }

    // Clean up
    document.body.removeChild(tempDiv);

    // Show success feedback
    setCopyButtonText('✓ Copied!');
    setShowSuccessCard(true);

    // Reset button text after 3 seconds
    setTimeout(() => {
      setCopyButtonText('Copy Signature');
    }, 3000);
  }
};

  // Render office address field with special handling
  const renderOfficeAddressField = (): JSX.Element => {
    const isEditing = editingField === 'officeAddress';
    const hasAddress = !!officeLocation;

    return (
      <div style={{ marginTop: '15px' }}>
        <div style={{ display: 'flex', alignItems: 'center', marginBottom: '5px' }}>
          <strong>Full Address:</strong>
          {isAddressCustomized && !isEditing && (
            <span style={{ 
              marginLeft: '10px', 
              fontSize: '8pt', 
              color: '#0078d4', 
              backgroundColor: '#e6f2ff', 
              padding: '2px 8px', 
              borderRadius: '3px' 
            }}>
              ✎ Custom
            </span>
          )}
          {hasAddress && !isEditing && (
            <div style={{ marginLeft: 'auto' }}>
              <IconButton
                iconProps={{ iconName: 'Edit' }}
                title="Edit"
                onClick={startEditOfficeAddress}
              />
            </div>
          )}
        </div>

        {isEditing ? (
          <div>
            <textarea
              value={tempValue}
              onChange={(e) => setTempValue(e.target.value)}
              style={{ 
                width: '100%', 
                minHeight: '60px',
                padding: '8px',
                border: '1px solid #ccc',
                borderRadius: '2px',
                fontFamily: 'Arial, sans-serif',
                fontSize: '9pt',
                resize: 'vertical'
              }}
              autoFocus
            />
            <div style={{ display: 'flex', gap: '5px', marginTop: '8px' }}>
              <IconButton
                iconProps={{ iconName: 'Refresh' }}
                title="Refresh (revert to dropdown selection)"
                onClick={refreshOfficeAddress}
                styles={{ root: { color: '#0078d4' } }}
              />
              <IconButton
                iconProps={{ iconName: 'Delete' }}
                title="Delete"
                onClick={deleteOfficeAddress}
                styles={{ root: { color: '#d13438' } }}
              />
              <IconButton
                iconProps={{ iconName: 'Accept' }}
                title="Save"
                onClick={saveOfficeAddress}
                styles={{ root: { color: '#107c10' } }}
              />
              <IconButton
                iconProps={{ iconName: 'Cancel' }}
                title="Cancel"
                onClick={cancelEdit}
                styles={{ root: { color: '#d13438' } }}
              />
            </div>
          </div>
        ) : (
          <div>
            <div style={{ 
              padding: '10px', 
              backgroundColor: '#f9f9f9', 
              borderRadius: '4px',
              border: '1px solid #e0e0e0',
              minHeight: '40px',
              fontSize: '9pt',
              color: hasAddress ? '#333' : '#999'
            }}>
              {hasAddress ? officeLocation : '(No location selected - please select office type above)'}
            </div>
            {hasAddress && (
              <div style={{ 
                fontSize: '8pt', 
                color: '#666', 
                marginTop: '8px',
                fontStyle: 'italic'
              }}>
                📍 This address will appear in your email signature
              </div>
            )}
          </div>
        )}
      </div>
    );
  };


// Render Quick Start Guide Card
const renderQuickStartGuide = (): JSX.Element | null => {
  if (!showQuickStart) {
    return (
      <div style={{ marginBottom: '15px' }}>
        <a 
          href="#" 
          onClick={(e) => { e.preventDefault(); setShowQuickStart(true); }}
          style={{ 
            color: '#0078d4', 
            textDecoration: 'none',
            fontSize: '9pt',
            display: 'flex',
            alignItems: 'center',
            gap: '5px'
          }}
        >
          📘 Show Quick Start Guide
        </a>
      </div>
    );
  }

  return (
    <div style={{
      border: '2px solid #0078d4',
      borderRadius: '4px',
      padding: '15px',
      backgroundColor: '#f0f6ff',
      marginBottom: '20px',
      position: 'relative'
    }}>
      {/* Close button */}
      <IconButton
        iconProps={{ iconName: 'Cancel' }}
        title="Close"
        onClick={() => setShowQuickStart(false)}
        styles={{
          root: {
            position: 'absolute',
            top: '10px',
            right: '10px',
            color: '#666'
          }
        }}
      />

      {/* Title */}
      <div style={{
        fontSize: '11pt',
        fontWeight: '600',
        color: '#0078d4',
        marginBottom: '12px',
        display: 'flex',
        alignItems: 'center',
        gap: '8px'
      }}>
        🎯 Quick Start Guide
      </div>

      {/* Content */}
      <div style={{ fontSize: '9pt', color: '#333', lineHeight: '1.6' }}>
        <p style={{ marginBottom: '10px' }}>
          Welcome! Generate your SEDC email signature in 4 easy steps:
        </p>
        
        <ol style={{ marginLeft: '20px', marginBottom: '10px' }}>
          <li style={{ marginBottom: '6px' }}>
            Click <strong>[Auto-fill]</strong> to load your SEDC profile
          </li>
          <li style={{ marginBottom: '6px' }}>
            <strong>Review, add, and edit</strong> your information as needed
          </li>
          <li style={{ marginBottom: '6px' }}>
            Select your <strong>office location</strong> from the dropdowns
          </li>
          <li style={{ marginBottom: '6px' }}>
            Click <strong>[Copy Signature]</strong> and paste in Outlook
          </li>
        </ol>

        <div style={{
          backgroundColor: '#e6f2ff',
          padding: '8px 10px',
          borderRadius: '3px',
          fontSize: '8pt',
          color: '#0078d4'
        }}>
          💡 First time? Click <strong>[? How to Use]</strong> in the top corner for detailed instructions
        </div>
      </div>
    </div>
  );
};

// Render Help Modal
const renderHelpModal = (): JSX.Element | null => {
  if (!showHelpModal) return null;

  return (
    <div style={{
      position: 'fixed',
      top: 0,
      left: 0,
      right: 0,
      bottom: 0,
      backgroundColor: 'rgba(0, 0, 0, 0.5)',
      zIndex: 1000,
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      padding: '20px'
    }}>
      <div style={{
        backgroundColor: 'white',
        borderRadius: '4px',
        maxWidth: '700px',
        maxHeight: '90vh',
        overflow: 'auto',
        boxShadow: '0 4px 16px rgba(0,0,0,0.2)'
      }}>
        {/* Header */}
        <div style={{
          padding: '20px',
          borderBottom: '1px solid #e0e0e0',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          position: 'sticky',
          top: 0,
          backgroundColor: 'white',
          zIndex: 1
        }}>
          <h2 style={{ margin: 0, fontSize: '14pt', color: '#333' }}>
            📖 How to Use SEDC Signature Generator
          </h2>
          <IconButton
            iconProps={{ iconName: 'Cancel' }}
            title="Close"
            onClick={() => setShowHelpModal(false)}
          />
        </div>

        {/* Content */}
        <div style={{ padding: '20px', fontSize: '9pt', lineHeight: '1.6' }}>
          
          {/* Getting Started */}
          <h3 style={{ fontSize: '11pt', color: '#0078d4', marginBottom: '10px' }}>
            Getting Started
          </h3>
          <hr style={{ border: 'none', borderTop: '1px solid #e0e0e0', marginBottom: '15px' }} />

          <div style={{ marginBottom: '20px' }}>
            <h4 style={{ fontSize: '10pt', marginBottom: '8px' }}>1️⃣ Auto-fill Your Information</h4>
            <p style={{ marginLeft: '20px', color: '#666' }}>
              Click <strong>[Auto-fill from SEDC Profile]</strong> to automatically load your name, email, and other details from your Microsoft account.
            </p>
          </div>

          <div style={{ marginBottom: '20px' }}>
            <h4 style={{ fontSize: '10pt', marginBottom: '8px' }}>2️⃣ Review, Add, and Edit Your Details</h4>
            <ul style={{ marginLeft: '35px', color: '#666' }}>
              <li>Click ✏️ next to any field to edit it</li>
              <li>Fields marked <span style={{ backgroundColor: '#fff4ce', color: '#856404', padding: '1px 4px', borderRadius: '2px', fontSize: '7pt' }}>Req</span> are required</li>
              <li>Use 🔄 to refresh from SEDC Profile</li>
              <li>Use 🗑️ to clear optional fields</li>
              <li>Add information like Personal Mobile manually if needed</li>
            </ul>
          </div>

          <div style={{ marginBottom: '20px' }}>
            <h4 style={{ fontSize: '10pt', marginBottom: '8px' }}>3️⃣ Select Your Office Location</h4>
            <ul style={{ marginLeft: '35px', color: '#666' }}>
              <li>Choose your office type from the dropdown</li>
              <li>Select specific location if you're in RO or PIBU</li>
              <li>You can customize the address after selecting</li>
            </ul>
          </div>

          <div style={{ marginBottom: '30px' }}>
            <h4 style={{ fontSize: '10pt', marginBottom: '8px' }}>4️⃣ Copy Your Signature</h4>
            <ul style={{ marginLeft: '35px', color: '#666' }}>
              <li>Click <strong>[Copy Signature]</strong> when all required fields are filled</li>
              <li>Your signature will be copied to clipboard</li>
              <li>Follow the instructions below to paste it in Outlook</li>
            </ul>
          </div>

          {/* Adding to Outlook */}
          <h3 style={{ fontSize: '11pt', color: '#0078d4', marginBottom: '10px' }}>
            📧 Adding Signature to Outlook
          </h3>
          <hr style={{ border: 'none', borderTop: '1px solid #e0e0e0', marginBottom: '15px' }} />

          {/* Outlook Web */}
          <div style={{
            backgroundColor: '#f9f9f9',
            border: '1px solid #e0e0e0',
            borderRadius: '4px',
            padding: '15px',
            marginBottom: '15px'
          }}>
            <h4 style={{ fontSize: '10pt', marginBottom: '10px', color: '#333' }}>
              Outlook Web (outlook.office.com)
            </h4>
            <ol style={{ marginLeft: '20px', color: '#666' }}>
              <li>Click ⚙️ <strong>Settings</strong> icon (top-right corner)</li>
              <li>Go to <strong>Account → Signatures</strong></li>
              <li>Click <strong>"+ New Signature"</strong> button</li>
              <li>Enter signature name (e.g., "SEDC Email")</li>
              <li>Paste your signature (Ctrl+V or Cmd+V)</li>
              <li>Click <strong>Save</strong></li>
            </ol>
          </div>

          {/* Outlook Desktop */}
          <div style={{
            backgroundColor: '#f9f9f9',
            border: '1px solid #e0e0e0',
            borderRadius: '4px',
            padding: '15px',
            marginBottom: '20px'
          }}>
            <h4 style={{ fontSize: '10pt', marginBottom: '10px', color: '#333' }}>
              Outlook Desktop (Windows/Mac)
            </h4>
            <ol style={{ marginLeft: '20px', color: '#666' }}>
              <li>Click <strong>File → Options → Mail</strong></li>
              <li>Click <strong>"Signatures..."</strong> button</li>
              <li>Click <strong>"New"</strong> to create a signature</li>
              <li>Enter signature name</li>
              <li>Paste signature in the edit box</li>
              <li>Click <strong>OK</strong> to save</li>
            </ol>
          </div>

          {/* Tips & Troubleshooting */}
          <h3 style={{ fontSize: '11pt', color: '#0078d4', marginBottom: '10px' }}>
            💡 Tips & Troubleshooting
          </h3>
          <hr style={{ border: 'none', borderTop: '1px solid #e0e0e0', marginBottom: '15px' }} />

          <ul style={{ marginLeft: '20px', color: '#666', marginBottom: '15px' }}>
            <li style={{ marginBottom: '8px' }}>
              If the logo doesn't appear, try pasting in Outlook Web first, then copying from there to Outlook Desktop
            </li>
            <li style={{ marginBottom: '8px' }}>
              Make sure all <span style={{ backgroundColor: '#fff4ce', color: '#856404', padding: '1px 4px', borderRadius: '2px', fontSize: '7pt' }}>Req</span> fields are filled before copying
            </li>
            <li style={{ marginBottom: '8px' }}>
              You can return to this generator anytime to update your signature
            </li>
            <li style={{ marginBottom: '8px' }}>
              Need assistance? Visit <a href="https://supportgo.sedc.my" target="_blank" style={{ color: '#0078d4', textDecoration: 'none' }}>SupportGo</a> for IT support
            </li>
          </ul>

        </div>

        {/* Footer */}
        <div style={{
          padding: '15px 20px',
          borderTop: '1px solid #e0e0e0',
          display: 'flex',
          justifyContent: 'flex-end',
          position: 'sticky',
          bottom: 0,
          backgroundColor: 'white'
        }}>
          <PrimaryButton
            text="Close"
            onClick={() => setShowHelpModal(false)}
          />
        </div>
      </div>
    </div>
  );
};

// Render Success Card
const renderSuccessCard = (): JSX.Element | null => {
  if (!showSuccessCard) return null;

  return (
    <div style={{
      border: '2px solid #107c10',
      borderRadius: '4px',
      padding: '15px',
      backgroundColor: '#f0fff0',
      marginTop: '15px',
      position: 'relative'
    }}>
      {/* Close button */}
      <IconButton
        iconProps={{ iconName: 'Cancel' }}
        title="Close"
        onClick={() => setShowSuccessCard(false)}
        styles={{
          root: {
            position: 'absolute',
            top: '10px',
            right: '10px',
            color: '#666'
          }
        }}
      />

      {/* Title */}
      <div style={{
        fontSize: '11pt',
        fontWeight: '600',
        color: '#107c10',
        marginBottom: '10px',
        display: 'flex',
        alignItems: 'center',
        gap: '8px'
      }}>
        ✓ Signature Copied Successfully!
      </div>

      {/* Content */}
      <div style={{ fontSize: '9pt', color: '#333', lineHeight: '1.6', marginRight: '30px' }}>
        <p style={{ marginBottom: '10px' }}>
          📋 Your signature is now in your clipboard
        </p>
        
        <p style={{ marginBottom: '8px', fontWeight: '600' }}>Next Steps:</p>
        <ol style={{ marginLeft: '20px', marginBottom: '10px' }}>
          <li>Open Outlook (Web or Desktop)</li>
          <li>Go to Signature settings</li>
          <li>Create new signature and paste (Ctrl+V)</li>
        </ol>

        <div style={{
          backgroundColor: '#e6f2ff',
          padding: '8px 10px',
          borderRadius: '3px',
          fontSize: '8pt',
          color: '#0078d4'
        }}>
          Need detailed instructions? → Click <strong>[? How to Use]</strong> button above
        </div>
      </div>
    </div>
  );
};

// Render editable field inline (label and field on same row)
const renderEditableFieldInline = (label: string, fieldName: string, value: string, isRequired: boolean = false): JSX.Element => {
  const isEditing = editingField === fieldName;
  const hasError = fieldErrors[fieldName];
  const canRetrieve = canRetrieveFromAAD(fieldName);
  const canDelete = !isRequired;

  return (
    <div style={{ marginBottom: '10px' }}>
      {/* Label and Field on same row */}
      <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
        {/* Label with required badge */}
        <div style={{ 
          minWidth: '130px', 
          display: 'flex', 
          alignItems: 'center',
          fontSize: '9pt',
          fontWeight: '500',
          color: '#333'
        }}>
          <span>{label}:</span>
          {isRequired && (
            <span style={{ 
              marginLeft: '6px', 
              fontSize: '7pt', 
              color: '#856404', 
              backgroundColor: '#fff4ce', 
              padding: '1px 4px', 
              borderRadius: '2px',
              fontWeight: '500'
            }}>
              Req
            </span>
          )}
        </div>

        {/* Field */}
        {isEditing ? (
          <div style={{ flex: 1, display: 'flex', alignItems: 'center', gap: '5px' }}>
            <input
              type="text"
              value={tempValue}
              onChange={(e) => setTempValue(e.target.value)}
              style={{ 
                flex: 1,
                padding: '6px 8px',
                border: '1px solid #ccc',
                borderRadius: '2px',
                fontSize: '9pt'
              }}
              autoFocus
            />
            {canRetrieve && (
              <IconButton
                iconProps={{ iconName: 'Refresh' }}
                title="Refresh from AAD"
                onClick={() => autoRetrieveFromAAD(fieldName)}
                styles={{ root: { color: '#0078d4', minWidth: '32px', height: '32px' } }}
              />
            )}
            {canDelete && (
              <IconButton
                iconProps={{ iconName: 'Delete' }}
                title="Delete"
                onClick={() => deleteField(fieldName)}
                styles={{ root: { color: '#d13438', minWidth: '32px', height: '32px' } }}
              />
            )}
            <IconButton
              iconProps={{ iconName: 'Accept' }}
              title="Save"
              onClick={() => saveEdit(fieldName)}
              styles={{ root: { color: '#107c10', minWidth: '32px', height: '32px' } }}
            />
            <IconButton
              iconProps={{ iconName: 'Cancel' }}
              title="Cancel"
              onClick={cancelEdit}
              styles={{ root: { color: '#d13438', minWidth: '32px', height: '32px' } }}
            />
          </div>
        ) : (
          <div style={{ flex: 1, display: 'flex', alignItems: 'center' }}>
            <div style={{ 
              flex: 1,
              padding: '6px 8px', 
              backgroundColor: '#f9f9f9', 
              borderRadius: '2px',
              border: '1px solid #e0e0e0',
              fontSize: '9pt',
              color: value ? '#333' : '#999'
            }}>
              {value || '(Not set)'}
            </div>
            <IconButton
              iconProps={{ iconName: 'Edit' }}
              title="Edit"
              onClick={() => startEdit(fieldName, value)}
              styles={{ root: { minWidth: '32px', height: '32px', marginLeft: '5px' } }}
            />
          </div>
        )}
      </div>

      {/* Error message below if exists */}
      {hasError && (
        <div style={{ 
          color: '#d13438', 
          fontSize: '8pt', 
          marginTop: '4px',
          marginLeft: '140px'
        }}>
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

      {/* ===== Top Bar ===== */}
      <div
        style={{
          display: 'flex',
          justifyContent: 'space-between',
          alignItems: 'center',
          marginBottom: '20px',
          paddingBottom: '15px',
          borderBottom: '2px solid #00a651'
        }}
      >
        <h2 style={{ margin: 0, fontSize: '14pt', color: '#333' }}>
          SEDC Email Signature Generator
        </h2>
        <IconButton
          iconProps={{ iconName: 'Help' }}
          title="How to Use"
          onClick={() => setShowHelpModal(true)}
          styles={{
            root: { fontSize: '12pt', color: '#0078d4' },
            label: { fontSize: '9pt', marginLeft: '5px' }
          }}
          text="How to Use"
        />
      </div>

      {/* ===== Auto-Fill Button ===== */}
      <div style={{ marginBottom: '20px' }}>
        <PrimaryButton
          text="🔄 Auto-fill from SEDC Profile"
          onClick={loadUserProfile}
          disabled={loading}
          iconProps={loading ? { iconName: 'ProgressRingDots' } : undefined}
          styles={{
            root: {
              width: '100%',
              height: '40px',
              backgroundColor: '#00a651',
              borderColor: '#00a651'
            },
            rootHovered: {
              backgroundColor: '#008a43',
              borderColor: '#008a43'
            }
          }}
        />
      </div>

      {/* ===== Quick Start Guide ===== */}
      {renderQuickStartGuide()}

      {/* ===== Error Message ===== */}
      {error && (
        <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setError(null)}>
          {error}
        </MessageBar>
      )}

      {/* ===== Loading Spinner ===== */}
      {loading && <Spinner label="Loading your profile..." />}

      {/* ===== Success Message ===== */}
      {copySuccess && (
        <MessageBar messageBarType={MessageBarType.success} isMultiline={false}>
          Signature copied to clipboard! Now paste it into Outlook settings.
        </MessageBar>
      )}

      {/* ===== Main Layout ===== */}
      {!loading && (
        <div className={styles.grid}>
          {/* ===== Left Panel (Form) ===== */}
          <div className={styles.controls}>
            {/* ==== User Information ==== */}
            <h3>User Information</h3>
            <div className={styles.section}>
              {/* Personal Details */}
              <div
                style={{
                  border: '1px solid #d0d0d0',
                  borderRadius: '4px',
                  padding: '15px',
                  backgroundColor: '#fafafa',
                  marginBottom: '15px'
                }}
              >
                <div
                  style={{
                    fontSize: '10pt',
                    fontWeight: '600',
                    marginBottom: '12px',
                    color: '#333'
                  }}
                >
                  Personal Details
                </div>

                {renderEditableFieldInline('Name', 'displayName', displayName, true)}
                {renderEditableFieldInline('Title', 'jobTitle', jobTitle)}
                {renderEditableFieldInline('Unit', 'unit', unit)}
                {renderEditableFieldInline('Department', 'department', department, true)}
              </div>

              {/* Contact Details */}
              <div
                style={{
                  border: '1px solid #d0d0d0',
                  borderRadius: '4px',
                  padding: '15px',
                  backgroundColor: '#fafafa',
                  marginBottom: '15px'
                }}
              >
                <div
                  style={{
                    fontSize: '10pt',
                    fontWeight: '600',
                    marginBottom: '12px',
                    color: '#333'
                  }}
                >
                  Contact Details
                </div>

                {renderEditableFieldInline('Email', 'mail', mail, true)}
                {renderEditableFieldInline('Business Phone', 'businessPhone', businessPhone)}
                {renderEditableFieldInline('Personal Mobile', 'personalMobile', personalMobile)}
              </div>

              {/* Info Message */}
              <div
                style={{
                  fontSize: '8pt',
                  color: '#0078d4',
                  backgroundColor: '#e6f2ff',
                  padding: '8px 10px',
                  borderRadius: '3px',
                  display: 'flex',
                  alignItems: 'center'
                }}
              >
                <span style={{ marginRight: '5px' }}>ℹ️</span>
                <span>
                  <span
                    style={{
                      backgroundColor: '#fff4ce',
                      color: '#856404',
                      padding: '1px 4px',
                      borderRadius: '2px',
                      fontSize: '7pt',
                      fontWeight: '500',
                      marginRight: '4px'
                    }}
                  >
                    Req
                  </span>
                  = Required field must be completed for signature
                </span>
              </div>
            </div>

            {/* ==== Office Location ==== */}
            <h3>Office Location</h3>
            <div className={styles.section}>
              {/* Office Selectors and Address */}
              {renderOfficeSelectors()}
              {renderOfficeAddressField()}
            </div>

            {/* ==== Copy Signature Button ==== */}
            <PrimaryButton
              text={copyButtonText}
              onClick={copySignature}
              iconProps={{ iconName: copyButtonText.includes('Copied') ? 'CheckMark' : 'Copy' }}
              disabled={!areRequiredFieldsFilled()}
              title={
                !areRequiredFieldsFilled()
                  ? `⚠️ Please complete all required fields:\n${getMissingFields()
                      .map((f) => `• ${f}`)
                      .join('\n')}`
                  : 'Copy signature to clipboard'
              }
              styles={{
                root: {
                  backgroundColor: copyButtonText.includes('Copied') ? '#107c10' : '#0078d4',
                  borderColor: copyButtonText.includes('Copied') ? '#107c10' : '#0078d4',
                  width: '100%',
                  marginTop: '20px'
                },
                rootDisabled: {
                  backgroundColor: '#f3f2f1',
                  borderColor: '#c8c6c4',
                  color: '#a19f9d'
                }
              }}
            />
          </div>

          {/* ===== Right Panel (Preview) ===== */}
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

            {renderSuccessCard()}
          </div>
        </div>
      )}

      {/* ===== Help Modal ===== */}
      {renderHelpModal()}
    </div>
  </div>
);

};

export default SedcSignatureGenerator;