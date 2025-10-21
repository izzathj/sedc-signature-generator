export interface OfficeAddress {
  type: 'simple' | 'nested';
  address?: string;
  locations?: Record<string, string>;
}

export const officeAddresses: Record<string, OfficeAddress> = {
  'Menara SEDC': {
    type: 'simple',
    address: 'Menara SEDC, No.2, The Isthmus, 93050, Kuching, Sarawak'
  },
  'SEDC Plaza': {
    type: 'simple',
    address: 'Level 8 & 9, SEDC Plaza, Jalan Menara SEDC, Tunku Abdul Rahman, No. 2, The Isthmus, 93100 Kuching, Sarawak'
  },
  
  'RO': {
    type: 'nested',
    locations: {
      'BETONG': 'Ground Floor, Sublot No. 34, Phase 3, Bandar Baru Betong, 95700, Betong, Sarawak',
      'SIBU': 'No. 40, Tingkat 1, Lot 2852, Jalan Intan, 96000 Sibu, Sarawak',
      'MUKAH': 'Lot 726, 1st Floor, Bangunan Mulajaya, Jalan Masjid, P.O. Box 169, 96400 Mukah, Sarawak',
      'BINTULU': 'Sublot 6N7, Bintulu Townsquare, Jalan Tun Ahmad Zaidi, 97000 Bintulu, Sarawak',
      'MIRI': 'Tingkat 1, Lot 6077, Pusat Bandar Shophouses, Desa Pujut, Bandar Baru, Permyjaya, 98107 Tudan Miri, Sarawak'
    }
  },
  'PIBU': {
    type: 'nested',
    locations: {
        'KUCHING': '1st & 2nd Flr, Sublot 12 & 13, Lot 14227 & 14228, Section 65 KTLD Metrocity New Township, Jalan Matang, 93048 Kuching, Sarawak',
        'KOTA SAMARAHAN': 'Lot No. 7933 (E & F), Blok 9, Tingkat 1, Muara Tuang Land District, Jalan Dato Mohd. Musa, 94300 Kota Samarahan, Sarawak',
        'LUNDU': 'Lot 70, Tingkat 2, Jalan Kubu, 94500 Lundu, Sarawak',
        'SERIAN': 'Tingkat 1 Sublot 11, Of Parent Lot 537, Block 9 Bukar Sadong, Land District 94700 Serian, Sarawak',
        'GEDONG': 'Anjung Usahawan Gedong, Kampung Gedong, 94800, Simunjan, Sarawak',
        'SIMUNJAN': 'Rural Transformation Centre (RTC), Kampung Taman Hijrah, 94800 Simunjan, Sarawak',
        'SEBUYAU': 'Rural Transformation Centre (RTC), Kampung Sampat, Sebangan, 94850 Sebuyau, Sarawak',
        'SRI AMAN': 'Tingkat 2, Lot 2010, Blok 3, Simanggang Town District, Jalan Kelab, 95000 Sri Aman, Sarawak',
        'BETONG': 'Tingkat 1, Sublot No. 34, Phase 3, Bandar Baru Betong, 95700 Betong, Sarawak',
        'SARATOK': '1st & 2nd Flr, Sublot 1, Lot 244 & 652, Saratok Town District, 95400 Saratok, Sarawak',
        'SARIKEI': 'Pusat Inkubator dan Bimbingan Usahawan (PIBU) Sarikei, 1st Floor Lot 2299, Block 36, Sarikei Land District, 96100 Sarikei, Sarawak',
        'DARO': 'Lot 26, Daro Land District, 96200 Daro, Sarawak',
        'SIBU': '12A, Lot 1732 & 1733, Tingkat 1, Jalan Kampung Datu, P.O. Box 470, 96000 Sibu, Sarawak',
        'MUKAH': 'Tingkat 1, Lot 725, Bangunan Mulajaya, Jalan Masjid, 96400 Mukah, Sarawak',
        'KAPIT': 'Tingkat 1, Lot 355, Jalan Kapit By Pass, Kapit Town District, 96800 Kapit, Sarawak',
        'SONG': 'Sublot 2, On Lot 365, Song, 96850 Kapit, Sarawak',
        'BINTULU': 'Sublot 6N7, Bintulu Townsquare, Jalan Tun Ahmad, Zaidi, 97000 Bintulu, Sarawak',
        'BELAGA': 'Ground Floor, Lot 1053, Block 2, Mamau Land District, 96900 Belaga, Sarawak',
        'BELURU': '1st Floor, Survey LOT 236 (Sublot 17), Off Parent LOT 21, Block 17, Bakong Land District, 98000 Miri, Sarawak',
        'LONG LAMA': 'Ground Floor Lot 142 Shophouse, Jalan Layang-Layang, 98300 Long Lama, Baram District, Miri, Sarawak',
        'MIRI': 'Tingkat 1, Lot 6077, Pusat Bandar Shophouses, Desa Pujut, Bandar Baru Permyjaya, 98107 Tudan, Miri, Sarawak',
        'MARUDI': 'Lot 315, Ground Floor, 109, Jalan Perpaduan, 98050 Marudi, Sarawak',
        'BARIO' : 'Pusat Inkubator dan Bimbingan Usahawan (PIBU) Bario, No. 13, Tingkat 1, Kampung Padang Pasir, Bario, 98050 Baram, Sarawak',
        'LIMBANG': 'Tingkat 1 - 2, Lot 2069, Bangunan Tabung Haji, Jalan Ricketts, 98700 Limbang, Sarawak',
        'LAWAS': 'Tingkat 1, Lot 490, Sublot 8, Jalan Pantai, 98850 Lawas, Sarawak'
    }
  }
};

// Helper to get office type display names
export const officeTypeLabels: Record<string, string> = {
  'Menara SEDC': 'Menara SEDC',
  'SEDC Plaza': 'SEDC Plaza',
  'RO': 'Regional Office (RO)',
  'PIBU': 'PIBU'
};

// Helper function to get final address
export const getFinalAddress = (officeType: string, specificLocation?: string): string => {
  const office = officeAddresses[officeType];
  
  if (!office) return '';
  
  if (office.type === 'simple') {
    return office.address || '';
  }
  
  if (office.type === 'nested' && specificLocation && office.locations) {
    return office.locations[specificLocation] || '';
  }
  
  return '';
};