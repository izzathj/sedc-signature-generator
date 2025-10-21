import { MSGraphClientV3 } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

export interface UserProfile {
  displayName: string;
  jobTitle: string;
  mail: string;
  mobilePhone: string;
  businessPhones: string[];
  officeLocation: string;
  department: string;
  companyName: string;
  photo?: string;
}

export class GraphService {
  private graphClient: MSGraphClientV3;

  constructor(graphClient: MSGraphClientV3) {
    this.graphClient = graphClient;
  }

  public async getUserProfile(): Promise<UserProfile> {
    try {
      const user: MicrosoftGraph.User = await this.graphClient
        .api('/me')
        .select('displayName,jobTitle,mail,mobilePhone,businessPhones,officeLocation,department,companyName')
        .get();

      let photoUrl: string | undefined;
      try {
        const photoBlob = await this.graphClient
          .api('/me/photo/$value')
          .get();
        photoUrl = URL.createObjectURL(photoBlob);
      } catch {
        console.log('No photo available or error fetching photo');
      }

      return {
        displayName: user.displayName || '',
        jobTitle: user.jobTitle || '',
        mail: user.mail || '',
        mobilePhone: user.mobilePhone || '',
        businessPhones: user.businessPhones || [],
        officeLocation: user.officeLocation || '',
        department: user.department || '',
        companyName: user.companyName || 'SEDC',
        photo: photoUrl
      };
    } catch (error) {
      console.error('Error fetching user profile:', error);
      throw error;
    }
  }
}