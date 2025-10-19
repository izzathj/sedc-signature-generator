import * as React from 'react';
import { useState, useEffect } from 'react';
import styles from './SedcSignatureGenerator.module.scss';
import type { ISedcSignatureGeneratorProps } from './ISedcSignatureGeneratorProps';
import { PrimaryButton, DefaultButton, Toggle, TextField, Spinner, MessageBar, MessageBarType } from '@fluentui/react';
import { GraphService, UserProfile } from '../../../services/GraphService';
import { SignatureTemplates, SignatureData } from '../../../templates/SignatureTemplates';

const SedcSignatureGenerator: React.FC<ISedcSignatureGeneratorProps> = (props) => {
  const [userProfile, setUserProfile] = useState<UserProfile | null>(null);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string>('');
  const [selectedTemplate, setSelectedTemplate] = useState<number>(1);
  const [includePhoto, setIncludePhoto] = useState<boolean>(true);
  //const [includeMobile, setIncludeMobile] = useState<boolean>(false);
  const [includeOffice, setIncludeOffice] = useState<boolean>(true);
  //const [customText, setCustomText] = useState<string>('');
  const [signatureHtml, setSignatureHtml] = useState<string>('');
  const [copySuccess, setCopySuccess] = useState<boolean>(false);
  const [personalMobile, setPersonalMobile] = useState<string>('');  
  const [unit, setUnit] = useState<string>('');
  const [personalLinkedIn, setPersonalLinkedIn] = useState<string>('');
  const [personalFacebook, setPersonalFacebook] = useState<string>('');
  const [personalInstagram, setPersonalInstagram] = useState<string>('');

  useEffect(() => {
    void loadUserProfile();
  }, []);

useEffect(() => {
  if (userProfile) {
    generateSignature();
  }
}, [userProfile, selectedTemplate, includePhoto, /*includeMobile,*/ includeOffice, unit, personalMobile, personalLinkedIn, personalFacebook, personalInstagram]);

  const loadUserProfile = async (): Promise<void> => {
    try {
      setLoading(true);
      const graphClient = await props.context.msGraphClientFactory.getClient('3');
      const graphService = new GraphService(graphClient);
      const profile = await graphService.getUserProfile();
      setUserProfile(profile);
      setError('');
    } catch (err) {
      setError('Failed to load user profile. Please ensure API permissions are granted.');
      console.error(err);
    } finally {
      setLoading(false);
    }
  };

  const generateSignature = (): void => {
    if (!userProfile) return;

    const data: SignatureData = {
      ...userProfile,
      includePhoto,
      //includeMobile,
      includeOffice,
      //customText: customText || undefined,
      unit: unit || undefined,
      personalMobile: personalMobile || undefined,
      personalLinkedIn: personalLinkedIn || undefined,
      personalFacebook: personalFacebook || undefined,
      personalInstagram: personalInstagram || undefined
    };

    let html = '';
    switch (selectedTemplate) {
      case 1:
        html = SignatureTemplates.generateTemplate1(data);
        break;
      case 2:
        html = SignatureTemplates.generateTemplate2(data);
        break;
      case 3:
        html = SignatureTemplates.generateTemplate3(data);
        break;
    case 4:  // Add this
        html = SignatureTemplates.generateTemplate4(data);
        break;
    }

    setSignatureHtml(html);
  };

  const copySignature = async (): Promise<void> => {
    try {
      // Create a temporary div to hold the HTML
      const tempDiv = document.createElement('div');
      tempDiv.innerHTML = signatureHtml;
      document.body.appendChild(tempDiv);

      // Select the content
      const range = document.createRange();
      range.selectNode(tempDiv);
      const selection = window.getSelection();
      if (selection) {
        selection.removeAllRanges();
        selection.addRange(range);
      }

      // Copy to clipboard
      document.execCommand('copy');

      // Clean up
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
            <h3>User Information</h3>
            <div className={styles.userInfo}>
              <p><strong>Name:</strong> {userProfile?.displayName}</p>
              <p><strong>Title:</strong> {userProfile?.jobTitle}</p>
              <p><strong>Department:</strong> {userProfile?.department}</p>
              <p><strong>Email:</strong> {userProfile?.mail}</p>
              {userProfile?.businessPhones && userProfile.businessPhones.length > 0 && (
                <p><strong>Phone:</strong> {userProfile.businessPhones[0]}</p>
              )}
            </div>

            <h3>Additional Information</h3>
            <TextField
              label="Unit"
              placeholder="e.g., Corporate Planning Unit"
              value={unit}
              onChange={(_, value) => setUnit(value || '')}
              description="Enter your unit name (not in Active Directory)"
            />
            <TextField
              label="Personal Mobile Phone"
              placeholder="e.g., +60 12-345 6789"
              value={personalMobile}
              onChange={(_, value) => setPersonalMobile(value || '')}
              description="Optional: Add your personal mobile number"
            />

            <h3>Personal Social Media (Optional)</h3>
            <TextField
              label="LinkedIn"
              placeholder="https://linkedin.com/in/yourname"
              value={personalLinkedIn}
              onChange={(_, value) => setPersonalLinkedIn(value || '')}
            />
            <TextField
              label="Facebook"
              placeholder="https://facebook.com/yourname"
              value={personalFacebook}
              onChange={(_, value) => setPersonalFacebook(value || '')}
            />
            <TextField
              label="Instagram"
              placeholder="https://instagram.com/yourname"
              value={personalInstagram}
              onChange={(_, value) => setPersonalInstagram(value || '')}
            />

            <h3>Template Selection</h3>
            <div className={styles.templateButtons}>
              <DefaultButton
                text="Professional"
                onClick={() => setSelectedTemplate(1)}
                primary={selectedTemplate === 1}
              />
              <DefaultButton
                text="Modern"
                onClick={() => setSelectedTemplate(2)}
                primary={selectedTemplate === 2}
              />
              <DefaultButton
                text="Minimal"
                onClick={() => setSelectedTemplate(3)}
                primary={selectedTemplate === 3}
              /> 
              <DefaultButton
                text="SEDC Clean"
                onClick={() => setSelectedTemplate(4)}
                primary={selectedTemplate === 4}
              />
            </div>

            <h3>Options</h3>
            <Toggle
              label="Include Profile Photo"
              checked={includePhoto}
              onChange={(_, checked) => setIncludePhoto(checked || false)}
            />
            {/* <Toggle
              label="Include Mobile Phone"
              checked={includeMobile}
              onChange={(_, checked) => setIncludeMobile(checked || false)}
            /> */}
            <Toggle
              label="Include Office Location"
              checked={includeOffice}
              onChange={(_, checked) => setIncludeOffice(checked || false)}
            />

            {/* <TextField
              label="Custom Text (Optional)"
              placeholder="e.g., Confidentiality notice, disclaimer..."
              multiline
              rows={3}
              value={customText}
              onChange={(_, value) => setCustomText(value || '')}
            /> */}

            <PrimaryButton
              text="Copy Signature"
              onClick={copySignature}
              iconProps={{ iconName: 'Copy' }}
              style={{ marginTop: '20px' }}
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
                <li>Click "Copy Signature" button above</li>
                <li>Open Outlook → File → Options → Mail → Signatures</li>
                <li>Click "New" to create a new signature</li>
                <li>Paste (Ctrl+V) into the signature editor</li>
                <li>Click "OK" to save</li>
              </ol>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default SedcSignatureGenerator;