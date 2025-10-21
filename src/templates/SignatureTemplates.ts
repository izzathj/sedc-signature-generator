// import { GraphService } from '../services/GraphService';

export interface SignatureData {
  displayName: string;
  jobTitle: string;
  department: string;
  companyName: string;
  mail: string;
  businessPhones: string[];
  mobilePhone: string;
  officeLocation: string;
  photo?: string;
  includePhoto: boolean;
  includeOffice: boolean;
  unit?: string;
  personalMobile?: string;
  personalLinkedIn?: string;
  personalFacebook?: string;
  personalInstagram?: string;
  personalTwitter?: string;      // ADD THIS
  personalTikTok?: string;        // ADD THIS
}

export class SignatureTemplates {
  private static calculateLogoSettings(data: SignatureData): { height: number; clipPercent: number } {
    // Count contact lines
    let lines = 0;
    
    // Always present (minimum 3)
    lines++; // Name
    lines++; // Department
    lines++; // Email
    
    // Conditional fields
    if (data.jobTitle) lines++;
    if (data.unit) lines++;
    if (data.businessPhones && data.businessPhones.length > 0) lines++;
    if (data.personalMobile) lines++;
    if (data.officeLocation && data.includeOffice) lines++;
    
    // Height and clip mapping based on lines
    const settings: Record<number, { height: number; clipPercent: number }> = {
      3: { height: 50, clipPercent: 25 },  // Minimal
      4: { height: 57, clipPercent: 21 },  // +1 field
      5: { height: 65, clipPercent: 17 },  // Standard
      6: { height: 75, clipPercent: 12 },  // +Mobile or Location
      7: { height: 90, clipPercent: 5 },   // Maximum (almost no crop)
      8: { height: 100, clipPercent: 2 },  // Edge case (all + extras)
    };
    
    // Return settings for the calculated lines (default to 3 if somehow less)
    return settings[lines] || settings[3];
  }

  public static generateTemplate1(data: SignatureData): string {
    const phone = data.businessPhones && data.businessPhones.length > 0 ? data.businessPhones[0] : '';
    const mobile = data.personalMobile || '';
    
    return `
      <table style="font-family: Arial, sans-serif; font-size: 10pt; color: #333; border-collapse: collapse;" cellpadding="0" cellspacing="0">
        <tr>
          ${data.includePhoto && data.photo ? `
            <td style="padding-right: 15px; vertical-align: top;">
              <img src="${data.photo}" alt="Profile" style="width: 80px; height: 80px; border-radius: 50%; object-fit: cover;" />
            </td>
          ` : ''}
          <td style="border-left: 3px solid #0078d4; padding-left: 15px; vertical-align: top;">
            <div style="font-size: 14pt; font-weight: bold; color: #0078d4; margin-bottom: 5px;">
              ${data.displayName}
            </div>
            <div style="font-size: 10pt; color: #666; margin-bottom: 3px;">
              ${data.jobTitle}
            </div>
            <div style="font-size: 10pt; color: #666; margin-bottom: 10px;">
              ${data.department} | ${data.companyName}
            </div>
            ${phone ? `
              <div style="margin-bottom: 3px;">
                <span style="color: #0078d4;">📞</span> ${phone}
              </div>
            ` : ''}
            ${mobile ? `
              <div style="margin-bottom: 3px;">
                <span style="color: #0078d4;">📱</span> ${mobile}
              </div>
            ` : ''}
            <div style="margin-bottom: 3px;">
              <span style="color: #0078d4;">✉️</span> <a href="mailto:${data.mail}" style="color: #0078d4; text-decoration: none;">${data.mail}</a>
            </div>
            ${data.includeOffice && data.officeLocation ? `
              <div style="margin-bottom: 3px;">
                <span style="color: #0078d4;">📍</span> ${data.officeLocation}
              </div>
            ` : ''}

            <div style="margin-top: 10px; padding-top: 10px; border-top: 1px solid #ddd;">
            <div style="font-size: 9pt; color: #888; margin-bottom: 5px;">Follow SEDC:</div>
            <div style="margin-bottom: 8px;">
                <a href="https://www.facebook.com/sedcsarawak/" target="_blank" style="margin-right: 8px; text-decoration: none;">
                <img src="https://cdn-icons-png.flaticon.com/512/733/733547.png" alt="Facebook" style="width: 20px; height: 20px; vertical-align: middle;" />
                </a>
                <a href="https://www.instagram.com/sedcsarawak/" target="_blank" style="text-decoration: none;">
                <img src="https://cdn-icons-png.flaticon.com/512/2111/2111463.png" alt="Instagram" style="width: 20px; height: 20px; vertical-align: middle;" />
                </a>
            </div>
            ${data.personalLinkedIn || data.personalFacebook || data.personalInstagram ? `
                <div style="font-size: 9pt; color: #888; margin-bottom: 5px; margin-top: 8px;">Connect with me:</div>
                <div>
                ${data.personalLinkedIn ? `
                    <a href="${data.personalLinkedIn}" target="_blank" style="margin-right: 8px; text-decoration: none;">
                    <img src="https://cdn-icons-png.flaticon.com/512/174/174857.png" alt="LinkedIn" style="width: 20px; height: 20px; vertical-align: middle;" />
                    </a>
                ` : ''}
                ${data.personalFacebook ? `
                    <a href="${data.personalFacebook}" target="_blank" style="margin-right: 8px; text-decoration: none;">
                    <img src="https://cdn-icons-png.flaticon.com/512/733/733547.png" alt="Facebook" style="width: 20px; height: 20px; vertical-align: middle;" />
                    </a>
                ` : ''}
                ${data.personalInstagram ? `
                    <a href="${data.personalInstagram}" target="_blank" style="text-decoration: none;">
                    <img src="https://cdn-icons-png.flaticon.com/512/2111/2111463.png" alt="Instagram" style="width: 20px; height: 20px; vertical-align: middle;" />
                    </a>
                ` : ''}
                </div>
            ` : ''}
            </div>


          </td>
        </tr>
      </table>
    `;
  }

  public static generateTemplate2(data: SignatureData): string {
    const phone = data.businessPhones && data.businessPhones.length > 0 ? data.businessPhones[0] : '';
    const mobile = data.personalMobile || '';
    
    return `
      <table style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; font-size: 10pt; color: #333; max-width: 500px;" cellpadding="0" cellspacing="0">
        <tr>
          <td style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); padding: 20px; color: white;">
            ${data.includePhoto && data.photo ? `
              <img src="${data.photo}" alt="Profile" style="width: 70px; height: 70px; border-radius: 50%; border: 3px solid white; margin-bottom: 10px;" />
            ` : ''}
            <div style="font-size: 16pt; font-weight: bold; margin-bottom: 5px;">
              ${data.displayName}
            </div>
            <div style="font-size: 11pt; opacity: 0.9;">
              ${data.jobTitle}
            </div>
          </td>
        </tr>
        <tr>
          <td style="padding: 15px; background: #f8f9fa;">
            <div style="margin-bottom: 5px;">
              <strong>${data.department}</strong> | ${data.companyName}
            </div>
            ${phone ? `<div style="margin-bottom: 3px;">☎️ ${phone}</div>` : ''}
            ${mobile ? `<div style="margin-bottom: 3px;">📱 ${mobile}</div>` : ''}
            <div style="margin-bottom: 3px;">
              ✉️ <a href="mailto:${data.mail}" style="color: #667eea; text-decoration: none;">${data.mail}</a>
            </div>

            <div style="margin-top: 10px; padding-top: 10px; border-top: 1px solid #ddd;">
            <div style="font-size: 9pt; color: #666; margin-bottom: 5px;">Follow SEDC:</div>
            <div style="margin-bottom: 8px;">
                <a href="https://www.facebook.com/sedcsarawak/" target="_blank" style="margin-right: 8px; text-decoration: none;">
                <img src="https://cdn-icons-png.flaticon.com/512/733/733547.png" alt="Facebook" style="width: 20px; height: 20px; vertical-align: middle;" />
                </a>
                <a href="https://www.instagram.com/sedcsarawak/" target="_blank" style="text-decoration: none;">
                <img src="https://cdn-icons-png.flaticon.com/512/2111/2111463.png" alt="Instagram" style="width: 20px; height: 20px; vertical-align: middle;" />
                </a>
            </div>
            ${data.personalLinkedIn || data.personalFacebook || data.personalInstagram ? `
                <div style="font-size: 9pt; color: #666; margin-bottom: 5px; margin-top: 8px;">Connect with me:</div>
                <div>
                ${data.personalLinkedIn ? `
                    <a href="${data.personalLinkedIn}" target="_blank" style="margin-right: 8px; text-decoration: none;">
                    <img src="https://cdn-icons-png.flaticon.com/512/174/174857.png" alt="LinkedIn" style="width: 20px; height: 20px; vertical-align: middle;" />
                    </a>
                ` : ''}
                ${data.personalFacebook ? `
                    <a href="${data.personalFacebook}" target="_blank" style="margin-right: 8px; text-decoration: none;">
                    <img src="https://cdn-icons-png.flaticon.com/512/733/733547.png" alt="Facebook" style="width: 20px; height: 20px; vertical-align: middle;" />
                    </a>
                ` : ''}
                ${data.personalInstagram ? `
                    <a href="${data.personalInstagram}" target="_blank" style="text-decoration: none;">
                    <img src="https://cdn-icons-png.flaticon.com/512/2111/2111463.png" alt="Instagram" style="width: 20px; height: 20px; vertical-align: middle;" />
                    </a>
                ` : ''}
                </div>
            ` : ''}
            </div>

            ${data.includeOffice && data.officeLocation ? `<div>📍 ${data.officeLocation}</div>` : ''}
 
          </td>
        </tr>
      </table>
    `;
  }

  public static generateTemplate3(data: SignatureData): string {
    const phone = data.businessPhones && data.businessPhones.length > 0 ? data.businessPhones[0] : '';
    const mobile = data.personalMobile || '';
    
    return `
      <table style="font-family: 'Courier New', monospace; font-size: 10pt; color: #333; border: 2px solid #333; padding: 15px;" cellpadding="0" cellspacing="0">
        <tr>
          <td>
            ${data.includePhoto && data.photo ? `
              <img src="${data.photo}" alt="Profile" style="width: 60px; height: 60px; border: 2px solid #333; margin-bottom: 10px;" />
            ` : ''}
            <div style="font-size: 14pt; font-weight: bold; margin-bottom: 5px; text-transform: uppercase;">
              ${data.displayName}
            </div>
            <div style="font-size: 10pt; margin-bottom: 3px;">
              ${data.jobTitle}
            </div>
            <div style="font-size: 10pt; margin-bottom: 10px;">
              ${data.department} / ${data.companyName}
            </div>
            <div style="border-top: 1px solid #333; padding-top: 10px;">
              ${phone ? `<div style="margin-bottom: 3px;">T: ${phone}</div>` : ''}
              ${mobile ? `<div style="margin-bottom: 3px;">M: ${mobile}</div>` : ''}
              <div style="margin-bottom: 3px;">E: <a href="mailto:${data.mail}" style="color: #333;">${data.mail}</a></div>
              ${data.includeOffice && data.officeLocation ? `<div>A: ${data.officeLocation}</div>` : ''}
            </div>

            <div style="margin-top: 10px; padding-top: 10px; border-top: 1px solid #333;">
            <div style="font-size: 9pt; margin-bottom: 5px;">SEDC:</div>
            <div style="margin-bottom: 8px;">
                <a href="https://www.facebook.com/sedcsarawak/" target="_blank" style="margin-right: 8px;">
                <img src="https://cdn-icons-png.flaticon.com/512/733/733547.png" alt="FB" style="width: 18px; height: 18px;" />
                </a>
                <a href="https://www.instagram.com/sedcsarawak/" target="_blank">
                <img src="https://cdn-icons-png.flaticon.com/512/2111/2111463.png" alt="IG" style="width: 18px; height: 18px;" />
                </a>
            </div>
            ${data.personalLinkedIn || data.personalFacebook || data.personalInstagram ? `
                <div style="font-size: 9pt; margin-bottom: 5px; margin-top: 8px;">Personal:</div>
                <div>
                ${data.personalLinkedIn ? `
                    <a href="${data.personalLinkedIn}" target="_blank" style="margin-right: 8px;">
                    <img src="https://cdn-icons-png.flaticon.com/512/174/174857.png" alt="LI" style="width: 18px; height: 18px;" />
                    </a>
                ` : ''}
                ${data.personalFacebook ? `
                    <a href="${data.personalFacebook}" target="_blank" style="margin-right: 8px;">
                    <img src="https://cdn-icons-png.flaticon.com/512/733/733547.png" alt="FB" style="width: 18px; height: 18px;" />
                    </a>
                ` : ''}
                ${data.personalInstagram ? `
                    <a href="${data.personalInstagram}" target="_blank">
                    <img src="https://cdn-icons-png.flaticon.com/512/2111/2111463.png" alt="IG" style="width: 18px; height: 18px;" />
                    </a>
                ` : ''}
                </div>
            ` : ''}
            </div>


          </td>
        </tr>
      </table>
    `;
  }

public static generateTemplate4(data: SignatureData): string {
  const phone = data.businessPhones && data.businessPhones.length > 0 ? data.businessPhones[0] : '';
  const mobile = data.personalMobile || '';
  const hasPersonalSocials = data.personalLinkedIn || data.personalFacebook || data.personalInstagram || data.personalTwitter || data.personalTikTok;
  const hasOfficeLocation = data.includeOffice && data.officeLocation;
  
  const { height: logoHeight, clipPercent } = this.calculateLogoSettings(data);

  return `
    <table cellpadding="0" cellspacing="0" border="0" style="font-family: Arial, sans-serif; font-size: 10pt; color: #333333; border-collapse: collapse; width: 600px; table-layout: fixed;">
      <tr>
        <td style="width: 160px; padding-right: 20px; border-right: 3px solid #00a651; vertical-align: ${hasPersonalSocials ? 'middle' : 'top'};">
          ${!hasPersonalSocials ? `
            <!-- Single cell layout when no personal socials -->
            <table cellpadding="0" cellspacing="0" border="0" style="width: 100%;">
              <tr>
                <td style="text-align: center; vertical-align: middle;">
                  <!-- DYNAMIC Cropped logo container -->
                  <div style="width: 140px; height: ${logoHeight}px; overflow: hidden; margin: 0 auto; display: flex; align-items: center; justify-content: center;">
                    <img src="https://sedc.com.my/wp-content/uploads/2025/08/SEDC-new-logo-2025-scaled.png" alt="SEDC Logo" style="width: 140px; height: auto; object-fit: cover; clip-path: inset(${clipPercent}% 0 ${clipPercent}% 0); display: block;" />
                  </div>
                </td>
              </tr>
              <tr>
                <td style="padding-top: 10px;">
                  <!-- Separator -->
                  <div style="border-top: 1px solid #e0e0e0; margin-bottom: 10px;"></div>
                  
                  <!-- Follow SEDC -->
                  <div style="font-size: 9pt; color: #888888; line-height: 16px;">
                    <span style="display: inline-block; vertical-align: middle; margin-right: 5px;">Follow SEDC:</span>
                    <a href="https://www.facebook.com/sedcsarawak/" target="_blank" style="text-decoration: none; display: inline-block; vertical-align: middle; margin-right: 5px;">
                      <img src="https://cdn-icons-png.flaticon.com/512/733/733547.png" alt="Facebook" width="16" height="16" style="display: block;" />
                    </a>
                    <a href="https://www.instagram.com/sedcsarawak/" target="_blank" style="text-decoration: none; display: inline-block; vertical-align: middle;">
                      <img src="https://cdn-icons-png.flaticon.com/512/2111/2111463.png" alt="Instagram" width="16" height="16" style="display: block;" />
                    </a>
                  </div>
                </td>
              </tr>
            </table>
          ` : `
            <!-- Regular layout when personal socials exist -->
            <div style="text-align: center;">
              <img src="https://sedc.com.my/wp-content/uploads/2025/08/SEDC-new-logo-2025-scaled.png" alt="SEDC Logo" width="140" style="height: auto; display: block; margin: 0 auto;" />
            </div>
          `}
        </td>
        
        <!-- Right column continues... -->
        
        <td style="padding-left: 20px; vertical-align: top; min-width: 400px;">
          <!-- Name -->
          <div style="font-size: 11pt; font-weight: bold; color: #00a651; margin-bottom: 3px;">
            ${data.displayName}
          </div>
          
          <!-- Job Title -->
          <div style="font-size: 10pt; color: #666666; margin-bottom: 2px;">
            ${data.jobTitle}
          </div>
          
          <!-- Unit and Department -->
          <div style="font-size: 10pt; color: #999999; font-style: italic; margin-bottom: 10px;">
            ${data.unit ? data.unit + ', ' : ''}${data.department || 'Group Digital and Technology'}
          </div>
          
          <!-- Contact Info (ALL lines) -->
          <div style="font-size: 9pt; line-height: 16px;">
            <div style="margin-bottom: 3px;">
              📧 <a href="mailto:${data.mail}" style="color: #00a651; text-decoration: none;">${data.mail}</a>
            </div>
            ${phone ? `
              <div style="margin-bottom: 3px;">📞 ${phone}</div>
            ` : ''}
            ${mobile ? `
              <div style="margin-bottom: 3px;">📱 ${mobile}</div>
            ` : ''}
            ${hasOfficeLocation ? `
              <div style="margin-bottom: 3px;">📍 ${data.officeLocation}</div>
            ` : ''}
          </div>
        </td>
      </tr>
      
      ${hasPersonalSocials ? `
        <!-- Footer Row (separator + content in same row) -->
        <tr>
          <td style="padding-right: 20px; padding-top: 10px; border-right: 3px solid #00a651; vertical-align: top;">
            <!-- Separator above footer -->
            <div style="border-top: 1px solid #e0e0e0; margin-bottom: 10px;"></div>
            
            <!-- Follow SEDC -->
            <div style="font-size: 8.5pt; color: #888888; line-height: 20px;">
              <span style="display: inline-block; vertical-align: middle; margin-right: 5px;">Follow SEDC:</span>
              <a href="https://www.facebook.com/sedcsarawak/" target="_blank" style="text-decoration: none; display: inline-block; vertical-align: middle; margin-right: 5px;">
                <img src="https://cdn-icons-png.flaticon.com/512/733/733547.png" alt="Facebook" width="16" height="16" style="display: block;" />
              </a>
              <a href="https://www.instagram.com/sedcsarawak/" target="_blank" style="text-decoration: none; display: inline-block; vertical-align: middle;">
                <img src="https://cdn-icons-png.flaticon.com/512/2111/2111463.png" alt="Instagram" width="16" height="16" style="display: block;" />
              </a>
            </div>
          </td>
          
          <td style="padding-left: 20px; padding-top: 10px; vertical-align: top;">
            <!-- Separator above footer -->
            <div style="border-top: 1px solid #e0e0e0; margin-bottom: 10px;"></div>
            
            <!-- Connect with me -->
            <div style="font-size: 8.5pt; color: #888888; line-height: 20px;">
                <span style="display: inline-block; vertical-align: middle; margin-right: 5px;">Connect with me:</span>
                ${data.personalLinkedIn ? `
                  <a href="${data.personalLinkedIn}" target="_blank" style="text-decoration: none; display: inline-block; vertical-align: middle; margin-right: 5px;">
                    <img src="https://cdn-icons-png.flaticon.com/512/174/174857.png" alt="LinkedIn" width="16" height="16" style="display: block;" />
                  </a>
                ` : ''}
                ${data.personalFacebook ? `
                  <a href="${data.personalFacebook}" target="_blank" style="text-decoration: none; display: inline-block; vertical-align: middle; margin-right: 5px;">
                    <img src="https://cdn-icons-png.flaticon.com/512/733/733547.png" alt="Facebook" width="16" height="16" style="display: block;" />
                  </a>
                ` : ''}
                ${data.personalInstagram ? `
                  <a href="${data.personalInstagram}" target="_blank" style="text-decoration: none; display: inline-block; vertical-align: middle; margin-right: 5px;">
                    <img src="https://cdn-icons-png.flaticon.com/512/2111/2111463.png" alt="Instagram" width="16" height="16" style="display: block;" />
                  </a>
                ` : ''}
                ${data.personalTwitter ? `
                  <a href="${data.personalTwitter}" target="_blank" style="text-decoration: none; display: inline-block; vertical-align: middle; margin-right: 5px;">
                    <img src="https://cdn-icons-png.flaticon.com/512/5968/5968830.png" alt="Twitter" width="16" height="16" style="display: block;" />
                  </a>
                ` : ''}
                ${data.personalTikTok ? `
                  <a href="${data.personalTikTok}" target="_blank" style="text-decoration: none; display: inline-block; vertical-align: middle;">
                    <img src="https://cdn-icons-png.flaticon.com/512/5968/5968812.png" alt="TikTok" width="16" height="16" style="display: block;" />
                  </a>
                ` : ''}
            </div>
          </td>
        </tr>
      ` : ''}
    </table>
  `;
}
}