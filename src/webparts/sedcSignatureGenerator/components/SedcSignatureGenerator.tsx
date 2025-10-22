import * as React from 'react';
import { useState, useEffect } from 'react';
import styles from './SedcSignatureGenerator.module.scss';
import { ISedcSignatureGeneratorProps } from './ISedcSignatureGeneratorProps';
import { PrimaryButton, DefaultButton, IconButton, MessageBar, MessageBarType, Spinner } from '@fluentui/react';
import { GraphService, UserProfile } from '../../../services/GraphService';
import { SignatureTemplates } from '../../../templates/SignatureTemplates';
import { officeAddresses, officeTypeLabels, getFinalAddress } from '../../../data/OfficeAddresses';

export default function SedcSignatureGenerator(props: ISedcSignatureGeneratorProps): JSX.Element {
  // DEBUG LOGS
  console.log('🔍 SEDC Signature Generator - Component Loading');
  console.log('🔍 Props received:', props);
  console.log('🔍 Context received:', props.context);
  console.log('🔍 Context exists?', !!props.context);
  console.log('🔍 msGraphClientFactory exists?', !!props.context?.msGraphClientFactory);

// Field placeholders
const fieldPlaceholders: Record<string, string> = {
  displayName: 'e.g., John Doe',
  jobTitle: 'e.g., Senior Manager',
  unit: 'e.g., Network Infrastructure & Security Unit',
  department: 'e.g., Group Digital and Technology',
  mail: 'e.g., john.doe@sedc.my',
  businessPhone: 'e.g., 082551555 ext 123',
  personalMobile: 'e.g., +60 12-345 6789',
  functionalName: 'e.g., Leave Admin, HR Support, IT Helpdesk',
  sharedEmail: 'e.g., hr@sedc.my, support@sedc.my',
  sharedPhone: 'e.g., +6082-416918'
};
  const { context } = props;

  // State variables
  const [userProfile, setUserProfile] = useState<UserProfile | null>(null);
  const [loading, setLoading] = useState<boolean>(false);
  const [error, setError] = useState<string | null>(null);
  const [signatureHtml, setSignatureHtml] = useState<string>('');
  
  // Signature type
  const [signatureType, setSignatureType] = useState<'personal' | 'shared'>('personal');

  // Personal mailbox fields
  const [displayName, setDisplayName] = useState<string>('');
  const [jobTitle, setJobTitle] = useState<string>('');
  const [department, setDepartment] = useState<string>('');
  const [mail, setMail] = useState<string>('');
  const [businessPhone, setBusinessPhone] = useState<string>('');
  const [unit, setUnit] = useState<string>('');
  const [personalMobile, setPersonalMobile] = useState<string>('');

  // Shared mailbox fields
  const [functionalName, setFunctionalName] = useState<string>('');
  const [sharedEmail, setSharedEmail] = useState<string>('');
  const [sharedPhone, setSharedPhone] = useState<string>('');

  // Office location
  const [officeLocation, setOfficeLocation] = useState<string>('');
  const [officeType, setOfficeType] = useState<string>('');
  const [specificLocation, setSpecificLocation] = useState<string>('');
  const [isAddressCustomized, setIsAddressCustomized] = useState<boolean>(false);
  const [originalOfficeLocation, setOriginalOfficeLocation] = useState<string>('');

  // UI state
  const [editingField, setEditingField] = useState<string | null>(null);
  const [tempValue, setTempValue] = useState<string>('');
  const [fieldErrors, setFieldErrors] = useState<Record<string, string>>({});
  const [showQuickStart, setShowQuickStart] = useState<boolean>(true);
  const [showHelpModal, setShowHelpModal] = useState<boolean>(false);
  const [showSuccessCard, setShowSuccessCard] = useState<boolean>(false);
  const [copyButtonText, setCopyButtonText] = useState<string>('Copy Signature');


  
  // Load user profile
  const loadUserProfile = async (): Promise<void> => {
    try {
      setLoading(true);
      setError(null);
      const profile = await GraphService.getUserProfileFromContext(context);
      setUserProfile(profile);

      setDisplayName(profile.displayName || '');
      setJobTitle(profile.jobTitle || '');
      setDepartment(profile.department || '');
      setMail(profile.mail || '');
      setBusinessPhone(profile.businessPhones && profile.businessPhones.length > 0 ? profile.businessPhones[0] : '');
      
      setLoading(false);
      } catch (err) {
        setError(err instanceof Error ? err.message : 'Failed to load profile');
        setLoading(false);
      }
  };

  // Generate signature
  function generateSignature(): void {
    const data = {
      displayName: signatureType === 'personal' ? displayName : functionalName,
      jobTitle: signatureType === 'personal' ? jobTitle : '',
      department: department,
      companyName: 'SEDC',
      mail: signatureType === 'personal' ? mail : sharedEmail,
      businessPhones: signatureType === 'personal' 
        ? (businessPhone ? [businessPhone] : [])
        : (sharedPhone ? [sharedPhone] : []),
      mobilePhone: signatureType === 'personal' ? (personalMobile || '') : '',
      officeLocation: officeLocation || '',
      includePhoto: true,
      includeOffice: !!officeLocation,
      unit: signatureType === 'personal' ? (unit || undefined) : undefined,
      personalMobile: signatureType === 'personal' ? (personalMobile || undefined) : undefined
    };
    


    const html = SignatureTemplates.generateTemplate4(data);
    setSignatureHtml(html);
  }

  // Update officeLocation whenever officeType or specificLocation changes
  useEffect(() => {
    const address = getFinalAddress(officeType, specificLocation);
    setOfficeLocation(address);
    setOriginalOfficeLocation(address);
    setIsAddressCustomized(false);
  }, [officeType, specificLocation]);

  // Auto-fill department when switching to shared mailbox tab
  useEffect(() => {
    if (signatureType === 'shared' && userProfile && userProfile.department) {
      setDepartment(userProfile.department);
    }
  }, [signatureType, userProfile]);

  // Generate signature whenever relevant fields change
  useEffect(() => {
    if (userProfile || signatureType === 'shared') {
      generateSignature();
    }
  }, [displayName, jobTitle, department, mail, businessPhone, unit, personalMobile, 
      officeLocation, functionalName, sharedEmail, sharedPhone, signatureType]);
  // Auto-load profile on component mount
  useEffect(() => {
    loadUserProfile().catch((err) => console.error('Failed to load profile on mount:', err));
    generateSignature();
  }, []);


  // Validation
  const areRequiredFieldsFilled = (): boolean => {
    if (signatureType === 'personal') {
      return !!(displayName && department && mail);
    } else {
      return !!(functionalName && department && sharedEmail);
    }
  };

  const getMissingFields = (): string[] => {
    const missing: string[] = [];
    if (signatureType === 'personal') {
      if (!displayName) missing.push('Name');
      if (!department) missing.push('Department');
      if (!mail) missing.push('Email');
    } else {
      if (!functionalName) missing.push('Functional Name');
      if (!department) missing.push('Department');
      if (!sharedEmail) missing.push('Email');
    }
    return missing;
  };

  // Edit functions
  const startEdit = (fieldName: string, currentValue: string): void => {
    setEditingField(fieldName);
    setTempValue(currentValue);
  };

  const cancelEdit = (): void => {
    setEditingField(null);
    setTempValue('');
  };

  const saveEdit = (fieldName: string): void => {
    const requiredFields = ['displayName', 'department', 'mail', 'functionalName', 'sharedEmail'];
    
    if (tempValue.trim() === '' && requiredFields.indexOf(fieldName) !== -1) {
      setFieldErrors(prev => ({
        ...prev,
        [fieldName]: 'This field is required'
      }));
      return;
    }

    const newErrors = { ...fieldErrors };
    delete newErrors[fieldName];
    setFieldErrors(newErrors);

    if (fieldName === 'displayName') setDisplayName(tempValue);
    else if (fieldName === 'jobTitle') setJobTitle(tempValue);
    else if (fieldName === 'unit') setUnit(tempValue);
    else if (fieldName === 'department') setDepartment(tempValue);
    else if (fieldName === 'mail') setMail(tempValue);
    else if (fieldName === 'businessPhone') setBusinessPhone(tempValue);
    else if (fieldName === 'personalMobile') setPersonalMobile(tempValue);
    else if (fieldName === 'functionalName') setFunctionalName(tempValue);
    else if (fieldName === 'sharedEmail') setSharedEmail(tempValue);
    else if (fieldName === 'sharedPhone') setSharedPhone(tempValue);

    setEditingField(null);
    setTempValue('');
  };

  const deleteField = (fieldName: string): void => {
    if (fieldName === 'jobTitle') setJobTitle('');
    else if (fieldName === 'unit') setUnit('');
    else if (fieldName === 'businessPhone') setBusinessPhone('');
    else if (fieldName === 'personalMobile') setPersonalMobile('');
    else if (fieldName === 'sharedPhone') setSharedPhone('');
    
    setEditingField(null);
    setTempValue('');
  };

  const canRetrieveFromAAD = (fieldName: string): boolean => {
    if (signatureType === 'shared') {
      return false;
    }
    const retrievableFields = ['displayName', 'jobTitle', 'unit', 'department', 'mail', 'businessPhone'];
    return retrievableFields.indexOf(fieldName) !== -1;
  };

  const autoRetrieveFromAAD = async (fieldName: string): Promise<void> => {
    if (!userProfile) return;
    console.log("userProfile", JSON.stringify(userProfile));

    if (fieldName === 'displayName') setTempValue(userProfile.displayName || '');
    else if (fieldName === 'jobTitle') setTempValue(userProfile.jobTitle || '');
    else if (fieldName === 'unit') setTempValue(userProfile.unit || '');
    else if (fieldName === 'department') setTempValue(userProfile.department || '');
    else if (fieldName === 'mail') setTempValue(userProfile.mail || '');
    else if (fieldName === 'businessPhone') setTempValue(userProfile.businessPhones && userProfile.businessPhones.length > 0 ? userProfile.businessPhones[0] : '');
  };

  // Office address functions
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

  // Copy signature
  const copySignature = (): void => {
    if (signatureHtml) {
      const tempDiv = document.createElement('div');
      tempDiv.innerHTML = signatureHtml;
      document.body.appendChild(tempDiv);

      const range = document.createRange();
      range.selectNodeContents(tempDiv);
      const selection = window.getSelection();
      if (selection) {
        selection.removeAllRanges();
        selection.addRange(range);
        document.execCommand('copy');
        selection.removeAllRanges();
      }

      document.body.removeChild(tempDiv);

      setCopyButtonText('✓ Copied!');
      setShowSuccessCard(true);

      setTimeout(() => {
        setCopyButtonText('Copy Signature');
      }, 3000);
    }
  };

  // Render editable field inline
  const renderEditableFieldInline = (label: string, fieldName: string, value: string, isRequired: boolean = false): JSX.Element => {
    const isEditing = editingField === fieldName;
    const hasError = fieldErrors[fieldName];
    const canRetrieve = canRetrieveFromAAD(fieldName);
    const canDelete = !isRequired;
    const placeholder = fieldPlaceholders[fieldName] || '(Not set)';

    return (
      <div style={{ marginBottom: '10px' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
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
                  title="Refresh from SEDC Profile"
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
                color: value ? '#333' : '#999',
                fontStyle: value ? 'normal' : 'italic'
              }}>
                {value || placeholder}
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

  // Render office address field
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

  // Render Quick Start Guide
  const renderQuickStartGuide = (): JSX.Element => {
    return (
      <div style={{ marginBottom: '15px' }}>
        <div style={{
          border: '1px solid #00a651',
          borderRadius: '4px',
          backgroundColor: '#f0fff4'
        }}>
          <div
            onClick={() => setShowQuickStart(!showQuickStart)}
            style={{
              padding: '8px 12px',
              display: 'flex',
              justifyContent: 'space-between',
              alignItems: 'center',
              cursor: 'pointer',
              borderBottom: showQuickStart ? '1px solid #00a651' : 'none'
            }}
            onMouseEnter={(e) => {
              e.currentTarget.style.backgroundColor = '#e6f9ed';
            }}
            onMouseLeave={(e) => {
              e.currentTarget.style.backgroundColor = 'transparent';
            }}
          >
            <span style={{
              fontSize: '9pt',
              fontWeight: '500',
              color: '#00a651'
            }}>
              💡 Quick Tips (click to {showQuickStart ? 'hide' : 'show'})
            </span>
            
            <a
              href="#"
              onClick={(e) => {
                e.preventDefault();
                e.stopPropagation();
                setShowHelpModal(true);
              }}
              style={{
                fontSize: '8pt',
                color: '#00a651',
                textDecoration: 'none',
                padding: '2px 6px'
              }}
              onMouseEnter={(e) => {
                e.currentTarget.style.textDecoration = 'underline';
              }}
              onMouseLeave={(e) => {
                e.currentTarget.style.textDecoration = 'none';
              }}
            >
              [? Help]
            </a>
          </div>

          {showQuickStart && (
            <div style={{
              padding: '12px',
              fontSize: '9pt',
              color: '#333',
              lineHeight: '1.6'
            }}>
              <p style={{ margin: '0 0 8px 0', fontWeight: '500' }}>
                Generate your SEDC email signature in 4 steps:
              </p>
              
              <div style={{ marginLeft: '10px' }}>
                <div style={{ marginBottom: '5px' }}>
                  1. Click [Auto-fill] to load your SEDC profile
                </div>
                <div style={{ marginBottom: '5px' }}>
                  2. Review, add, and edit your information
                </div>
                <div style={{ marginBottom: '5px' }}>
                  3. Select your office location
                </div>
                <div style={{ marginBottom: '5px' }}>
                  4. Copy and paste into Outlook
                </div>
              </div>

              <div style={{
                marginTop: '10px',
                fontSize: '8pt',
                color: '#666',
                fontStyle: 'italic'
              }}>
                Need detailed instructions? Click [? Help] above
              </div>
            </div>
          )}
        </div>
      </div>
    );
  };

  // Render tabs
  const renderTabs = (): JSX.Element => {
    return (
      <div style={{
        display: 'flex',
        gap: '4px',
        marginBottom: '-1px',
        position: 'relative',
        zIndex: 1
      }}>
        <button
          onClick={() => setSignatureType('personal')}
          style={{
            padding: '8px 20px',
            fontSize: '10pt',
            fontWeight: signatureType === 'personal' ? '600' : '400',
            color: signatureType === 'personal' ? '#333' : '#666',
            backgroundColor: signatureType === 'personal' ? '#fafafa' : '#f5f5f5',
            border: '1px solid #d0d0d0',
            borderBottom: signatureType === 'personal' ? 'none' : '1px solid #d0d0d0',
            borderTopLeftRadius: '8px',
            borderTopRightRadius: '8px',
            cursor: signatureType === 'personal' ? 'default' : 'pointer',
            position: 'relative',
            zIndex: signatureType === 'personal' ? 3 : 1,
            transition: 'background-color 0.2s ease'
          }}
          onMouseEnter={(e) => {
            if (signatureType !== 'personal') {
              e.currentTarget.style.backgroundColor = '#ebebeb';
            }
          }}
          onMouseLeave={(e) => {
            if (signatureType !== 'personal') {
              e.currentTarget.style.backgroundColor = '#f5f5f5';
            }
          }}
        >
          Personal Mailbox
        </button>

        <button
          onClick={() => setSignatureType('shared')}
          style={{
            padding: '8px 20px',
            fontSize: '10pt',
            fontWeight: signatureType === 'shared' ? '600' : '400',
            color: signatureType === 'shared' ? '#333' : '#666',
            backgroundColor: signatureType === 'shared' ? '#fafafa' : '#f5f5f5',
            border: '1px solid #d0d0d0',
            borderBottom: signatureType === 'shared' ? 'none' : '1px solid #d0d0d0',
            borderTopLeftRadius: '8px',
            borderTopRightRadius: '8px',
            cursor: signatureType === 'shared' ? 'default' : 'pointer',
            position: 'relative',
            zIndex: signatureType === 'shared' ? 3 : 1,
            transition: 'background-color 0.2s ease'
          }}
          onMouseEnter={(e) => {
            if (signatureType !== 'shared') {
              e.currentTarget.style.backgroundColor = '#ebebeb';
            }
          }}
          onMouseLeave={(e) => {
            if (signatureType !== 'shared') {
              e.currentTarget.style.backgroundColor = '#f5f5f5';
            }
          }}
        >
          Shared Mailbox
        </button>
      </div>
    );
  };

  // Render Help Modal (placeholder - you'll need to add the full modal content)
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
              Click <strong>[Auto-fill]</strong> button to automatically load your name, email, and other available details from your SEDC profile.
            </p>
          </div>

          <div style={{ marginBottom: '20px' }}>
            <h4 style={{ fontSize: '10pt', marginBottom: '8px' }}>2️⃣ Review, Add, and Edit Your Details</h4>
            <ul style={{ color: '#666' }}>
              <li>Click ✏️ next to any field to edit it</li>
              <li>Fields marked <span style={{ backgroundColor: '#fff4ce', color: '#856404', padding: '1px 4px', borderRadius: '2px', fontSize: '7pt' }}>Req</span> are required</li>
              <li>Use 🔄 to refresh from SEDC Profile</li>
              <li>Use 🗑️ to clear optional fields</li>
              <li>Add information like Personal Mobile manually if needed</li>
            </ul>
          </div>

          <div style={{ marginBottom: '20px' }}>
            <h4 style={{ fontSize: '10pt', marginBottom: '8px' }}>3️⃣ Select Your Office Location</h4>
            <ul style={{ color: '#666' }}>
              <li>Choose your office type from the dropdown</li>
              <li>Select specific location if you&apos;re in RO or PIBU</li>
              <li>You can customize the address after selecting</li>
            </ul>
          </div>

          <div style={{ marginBottom: '30px' }}>
            <h4 style={{ fontSize: '10pt', marginBottom: '8px' }}>4️⃣ Copy Your Signature</h4>
            <ul style={{ color: '#666' }}>
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
            <ol style={{ color: '#666' }}>
              <li>Click ⚙️ <strong>Settings</strong> icon (top-right corner)</li>
              <li>Go to <strong>Account → Signatures</strong></li>
              <li>Click <strong>&quot;+ New Signature&quot;</strong> button</li>
              <li>Enter signature name (e.g., &quot;SEDC Email&quot;)</li>
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
            <ol style={{ color: '#666' }}>
              <li>Click <strong>File → Options → Mail</strong></li>
              <li>Click <strong>&quot;Signatures...&quot;</strong> button</li>
              <li>Click <strong>&quot;New&quot;</strong> to create a signature</li>
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
              If the logo doesn&apos;t appear, try pasting in Outlook Web first, then copying from there to Outlook Desktop
            </li>
            <li style={{ marginBottom: '8px' }}>
              Make sure all <span style={{ backgroundColor: '#fff4ce', color: '#856404', padding: '1px 4px', borderRadius: '2px', fontSize: '7pt' }}>Req</span> fields are filled before copying
            </li>
            <li style={{ marginBottom: '8px' }}>
              You can return to this generator anytime to update your signature
            </li>
            <li style={{ marginBottom: '8px' }}>
              Need assistance? Visit <a href="https://supportgo.sedc.my" target="_blank" rel="noopener noreferrer" style={{ color: '#0078d4', textDecoration: 'none' }}>SupportGo</a> for IT support
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

        <div style={{
          fontSize: '11pt',
          fontWeight: '600',
          color: '#107c10',
          marginBottom: '10px'
        }}>
          ✓ Signature Copied Successfully!
        </div>

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
            Need detailed instructions? → Click <strong>[? Help]</strong> button above
          </div>
        </div>
      </div>
    );
  };

  // Main return
  return (
    <div className={styles.sedcSignatureGenerator}>
      <div className={styles.container}>

      <div
        style={{
          marginBottom: '20px',
          paddingBottom: '15px',
          borderBottom: '2px solid #00a651',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between'
        }}
      >
        <h2 style={{ margin: 0, fontSize: '14pt', color: '#333' }}>
          SEDC Email Signature Generator
        </h2>
        <img 
          src="https://sedc.com.my/wp-content/uploads/2025/08/SEDC-new-logo-2025-scaled.png" 
          alt="SEDC Logo" 
          style={{ 
            height: '40px',
            width: 'auto'
          }} 
        />
      </div>
        {renderQuickStartGuide()}

        {error && (
          <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setError(null)}>
            {error}
          </MessageBar>
        )}

        {loading && <Spinner label="Loading your profile..." />}

        {!loading && (
          <div className={styles.grid}>
            <div className={styles.controls}>
              
              {renderTabs()}

              <div style={{
                border: '1px solid #d0d0d0',
                borderRadius: '0 4px 4px 4px',
                backgroundColor: '#fafafa',
                padding: '15px',
                position: 'relative',
                zIndex: 2
              }}>
                <div style={{ 
                  display: 'flex', 
                  justifyContent: 'space-between', 
                  alignItems: 'center',
                  marginBottom: '10px'
                }}>
                  <h3 style={{ margin: 0 }}>User Information</h3>
                  <DefaultButton
                    text="Auto-fill"
                    iconProps={{ iconName: 'Refresh' }}
                    onClick={loadUserProfile}
                    disabled={loading || signatureType === 'shared'}
                    styles={{
                      root: {
                        fontSize: '9pt',
                        padding: '4px 12px',
                        height: '28px',
                        border: '1px solid #00a651',
                        color: signatureType === 'shared' ? '#999' : '#00a651',
                        backgroundColor: signatureType === 'shared' ? '#f0f0f0' : 'transparent'
                      },
                      rootHovered: {
                        backgroundColor: signatureType === 'shared' ? '#f0f0f0' : '#f0fff0',
                        borderColor: signatureType === 'shared' ? '#d0d0d0' : '#008a43'
                      },
                      rootPressed: {
                        backgroundColor: '#e6f7e6'
                      },
                      rootDisabled: {
                        backgroundColor: '#f0f0f0',
                        borderColor: '#d0d0d0',
                        color: '#999'
                      },
                      icon: {
                        fontSize: '10pt'
                      }
                    }}
                    title={signatureType === 'shared' ? 'Auto-fill not available for shared mailboxes' : 'Auto-fill from SEDC Profile'}
                  />
                </div>

                {signatureType === 'personal' ? (
                  <>
                    <div style={{
                      border: '1px solid #d0d0d0',
                      borderRadius: '4px',
                      padding: '15px',
                      backgroundColor: '#fafafa',
                      marginBottom: '15px'
                    }}>
                      <div style={{
                        fontSize: '10pt',
                        fontWeight: '600',
                        marginBottom: '12px',
                        color: '#333'
                      }}>
                        Personal Details
                      </div>

                      {renderEditableFieldInline('Name', 'displayName', displayName, true)}
                      {renderEditableFieldInline('Title', 'jobTitle', jobTitle)}
                      {renderEditableFieldInline('Unit', 'unit', unit)}
                      {renderEditableFieldInline('Department', 'department', department, true)}
                    </div>

                    <div style={{
                      border: '1px solid #d0d0d0',
                      borderRadius: '4px',
                      padding: '15px',
                      backgroundColor: '#fafafa',
                      marginBottom: '15px'
                    }}>
                      <div style={{
                        fontSize: '10pt',
                        fontWeight: '600',
                        marginBottom: '12px',
                        color: '#333'
                      }}>
                        Contact Details
                      </div>

                      {renderEditableFieldInline('Email', 'mail', mail, true)}
                      {renderEditableFieldInline('Business Phone', 'businessPhone', businessPhone)}
                      {renderEditableFieldInline('Personal Mobile', 'personalMobile', personalMobile)}
                    </div>
                  </>
                ) : (
                  <>
                    <div style={{
                      fontSize: '8pt',
                      color: '#0078d4',
                      backgroundColor: '#e6f2ff',
                      padding: '8px 10px',
                      borderRadius: '3px',
                      marginBottom: '15px'
                    }}>
                      ℹ️ Shared mailbox signatures are for departmental/functional emails (e.g., hr@sedc.my, support@sedc.my)
                    </div>

                    <div style={{
                      border: '1px solid #d0d0d0',
                      borderRadius: '4px',
                      padding: '15px',
                      backgroundColor: '#fafafa',
                      marginBottom: '15px'
                    }}>
                      <div style={{
                        fontSize: '10pt',
                        fontWeight: '600',
                        marginBottom: '12px',
                        color: '#333'
                      }}>
                        Shared Mailbox Details
                      </div>

                      {renderEditableFieldInline('Functional Name', 'functionalName', functionalName, true)}
                      {renderEditableFieldInline('Department', 'department', department, true)}
                    </div>

                    <div style={{
                      border: '1px solid #d0d0d0',
                      borderRadius: '4px',
                      padding: '15px',
                      backgroundColor: '#fafafa',
                      marginBottom: '15px'
                    }}>
                      <div style={{
                        fontSize: '10pt',
                        fontWeight: '600',
                        marginBottom: '12px',
                        color: '#333'
                      }}>
                        Contact Details
                      </div>

                      {renderEditableFieldInline('Email', 'sharedEmail', sharedEmail, true)}
                      {renderEditableFieldInline('Business Phone', 'sharedPhone', sharedPhone)}
                    </div>
                  </>
                )}

                <div style={{
                  fontSize: '8pt',
                  color: '#0078d4',
                  backgroundColor: '#e6f2ff',
                  padding: '8px 10px',
                  borderRadius: '3px',
                  display: 'flex',
                  alignItems: 'center'
                }}>
                  <span style={{ marginRight: '5px' }}>ℹ️</span>
                  <span>
                    <span style={{
                      backgroundColor: '#fff4ce',
                      color: '#856404',
                      padding: '1px 4px',
                      borderRadius: '2px',
                      fontSize: '7pt',
                      fontWeight: '500',
                      marginRight: '4px'
                    }}>
                      Req
                    </span>
                    = Required field must be completed for signature
                  </span>
                </div>
              </div>

              <h3>Office Location</h3>
              <div className={styles.section}>
                <div style={{
                  border: '1px solid #d0d0d0',
                  borderRadius: '4px',
                  padding: '15px',
                  backgroundColor: '#fafafa',
                  marginBottom: '15px'
                }}>
                  <div style={{
                    fontSize: '10pt',
                    fontWeight: '600',
                    marginBottom: '12px',
                    color: '#333'
                  }}>
                    Select Your Office
                  </div>

                  <div style={{ marginBottom: '12px' }}>
                    <label style={{
                      fontSize: '9pt',
                      fontWeight: '500',
                      marginBottom: '5px',
                      display: 'block'
                    }}>
                      Office Type:
                    </label>
                    <select
                      value={officeType}
                      onChange={(e) => {
                        setOfficeType(e.target.value);
                        setSpecificLocation('');
                      }}
                      style={{
                        width: '100%',
                        padding: '8px',
                        border: '1px solid #ccc',
                        borderRadius: '2px',
                        fontSize: '9pt'
                      }}
                    >
                      <option value="">-- Select Office Type --</option>
                      {Object.keys(officeAddresses).map((key) => (
                        <option key={key} value={key}>
                          {officeTypeLabels[key]}
                        </option>
                      ))}
                    </select>
                    <div style={{
                      fontSize: '8pt',
                      color: '#666',
                      marginTop: '4px',
                      fontStyle: 'italic'
                    }}>
                      💡 Select your primary office location
                    </div>
                  </div>

                  {officeType === 'RO' && officeAddresses.RO.locations && (
                    <div style={{ marginBottom: '12px' }}>
                      <label style={{
                        fontSize: '9pt',
                        fontWeight: '500',
                        marginBottom: '5px',
                        display: 'block'
                      }}>
                        Specific Location:
                      </label>
                      <select
                        value={specificLocation}
                        onChange={(e) => setSpecificLocation(e.target.value)}
                        style={{
                          width: '100%',
                          padding: '8px',
                          border: '1px solid #ccc',
                          borderRadius: '2px',
                          fontSize: '9pt'
                        }}
                      >
                        <option value="">-- Select RO Location --</option>
                        {Object.keys(officeAddresses.RO.locations).map((key) => (
                          <option key={key} value={key}>
                            {key.charAt(0) + key.slice(1).toLowerCase()}
                          </option>
                        ))}
                      </select>
                      <div style={{
                        fontSize: '8pt',
                        color: '#666',
                        marginTop: '4px',
                        fontStyle: 'italic'
                      }}>
                        (Select your Regional Office location)
                      </div>
                    </div>
                  )}

                  {officeType === 'PIBU' && officeAddresses.PIBU.locations && (
                    <div style={{ marginBottom: '12px' }}>
                      <label style={{
                        fontSize: '9pt',
                        fontWeight: '500',
                        marginBottom: '5px',
                        display: 'block'
                      }}>
                        Specific Location:
                      </label>
                      <select
                        value={specificLocation}
                        onChange={(e) => setSpecificLocation(e.target.value)}
                        style={{
                          width: '100%',
                          padding: '8px',
                          border: '1px solid #ccc',
                          borderRadius: '2px',
                          fontSize: '9pt'
                        }}
                      >
                        <option value="">-- Select PIBU Location --</option>
                        {Object.keys(officeAddresses.PIBU.locations).map((key) => (
                          <option key={key} value={key}>
                            {key.charAt(0) + key.slice(1).toLowerCase()}
                          </option>
                        ))}
                      </select>
                      <div style={{
                        fontSize: '8pt',
                        color: '#666',
                        marginTop: '4px',
                        fontStyle: 'italic'
                      }}>
                        (Select your PIBU location)
                      </div>
                    </div>
                  )}

                  <div style={{
                    fontSize: '8pt',
                    color: '#0078d4',
                    backgroundColor: '#e6f2ff',
                    padding: '8px 10px',
                    borderRadius: '3px',
                    marginTop: '10px'
                  }}>
                    💡 Tip: You can edit the address below after selecting
                  </div>
                </div>

                {renderOfficeAddressField()}
              </div>

              <PrimaryButton
                text={copyButtonText}
                onClick={copySignature}
                iconProps={{ iconName: copyButtonText.indexOf('Copied') !== -1 ? 'CheckMark' : 'Copy' }}
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
                    backgroundColor: copyButtonText.indexOf('Copied') !== -1 ? '#107c10' : '#0078d4',
                    borderColor: copyButtonText.indexOf('Copied') !== -1 ? '#107c10' : '#0078d4',
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

            <div className={styles.preview}>
              <h3>Preview</h3>
              <div
                className={styles.previewBox}
                dangerouslySetInnerHTML={{ __html: signatureHtml }}
              />

              {renderSuccessCard()}
            </div>
          </div>
        )}

        {renderHelpModal()}
      </div>
    </div>
  );
}