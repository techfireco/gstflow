"""
GST Configuration Constants: State Codes, Standard Names, UQC, and Tax Rates.
"""

# 2-digit GST state code to Standard Name (as required in GSTR-1 Excel & JSON)
GST_STATE_CODES = {
    '01': '01-Jammu & Kashmir',
    '02': '02-Himachal Pradesh',
    '03': '03-Punjab',
    '04': '04-Chandigarh',
    '05': '05-Uttarakhand',
    '06': '06-Haryana',
    '07': '07-Delhi',
    '08': '08-Rajasthan',
    '09': '09-Uttar Pradesh',
    '10': '10-Bihar',
    '11': '11-Sikkim',
    '12': '12-Arunachal Pradesh',
    '13': '13-Nagaland',
    '14': '14-Manipur',
    '15': '15-Mizoram',
    '16': '16-Tripura',
    '17': '17-Meghalaya',
    '18': '18-Assam',
    '19': '19-West Bengal',
    '20': '20-Jharkhand',
    '21': '21-Odisha',
    '22': '22-Chhattisgarh',
    '23': '23-Madhya Pradesh',
    '24': '24-Gujarat',
    '26': '26-Dadra & Nagar Haveli & Daman & Diu',
    '27': '27-Maharashtra',
    '29': '29-Karnataka',
    '30': '30-Goa',
    '31': '31-Lakshadweep',
    '32': '32-Kerala',
    '33': '33-Tamil Nadu',
    '34': '34-Puducherry',
    '35': '35-Andaman & Nicobar Islands',
    '36': '36-Telangana',
    '37': '37-Andhra Pradesh',
    '38': '38-Ladakh',
    '97': '97-Other Territory'
}

# Fuzzy State Name matching dictionary (handles marketplace variants)
STATE_NAME_TO_CODE = {
    'JAMMU AND KASHMIR': '01',
    'JAMMU & KASHMIR': '01',
    'HIMACHAL PRADESH': '02',
    'PUNJAB': '03',
    'CHANDIGARH': '04',
    'UTTARAKHAND': '05',
    'UTTARANCHAL': '05',
    'HARYANA': '06',
    'DELHI': '07',
    'NEW DELHI': '07',
    'RAJASTHAN': '08',
    'UTTAR PRADESH': '09',
    'BIHAR': '10',
    'SIKKIM': '11',
    'ARUNACHAL PRADESH': '12',
    'NAGALAND': '13',
    'MANIPUR': '14',
    'MIZORAM': '15',
    'TRIPURA': '16',
    'MEGHALAYA': '17',
    'ASSAM': '18',
    'WEST BENGAL': '19',
    'JHARKHAND': '20',
    'ODISHA': '21',
    'ORISSA': '21',
    'CHHATTISGARH': '22',
    'CHATTISGARH': '22',
    'MADHYA PRADESH': '23',
    'GUJARAT': '24',
    'DADRA AND NAGAR HAVELI AND DAMAN AND DIU': '26',
    'DADRA & NAGAR HAVELI & DAMAN & DIU': '26',
    'DAMAN AND DIU': '26',
    'DADRA AND NAGAR HAVELI': '26',
    'MAHARASHTRA': '27',
    'KARNATAKA': '29',
    'GOA': '30',
    'LAKSHADWEEP': '31',
    'KERALA': '32',
    'TAMIL NADU': '33',
    'PUDUCHERRY': '34',
    'PONDICHERRY': '34',
    'ANDAMAN AND NICOBAR ISLANDS': '35',
    'ANDAMAN & NICOBAR': '35',
    'TELANGANA': '36',
    'ANDHRA PRADESH': '37',
    'ANDHRA PRADESH(NEW)': '37',
    'LADAKH': '38'
}

def normalize_state_code(raw_state: str) -> str:
    """Normalizes any state input string to a 2-digit GST state code."""
    if not raw_state:
        return '97'
    s = str(raw_state).strip().upper()
    # Check if starts with 2 digits (e.g. '07-DELHI' or '07')
    if len(s) >= 2 and s[:2].isdigit():
        code = s[:2]
        if code in GST_STATE_CODES:
            return code
    # Try exact lookup in name map
    if s in STATE_NAME_TO_CODE:
        return STATE_NAME_TO_CODE[s]
    # Fuzzy contains match
    for name, code in STATE_NAME_TO_CODE.items():
        if name in s or s in name:
            return code
    return '97'

# Known Marketplace TCS ETINs
ECO_GSTINS = {
    'MEESHO': {
        'default': '07AARCM9332R1CQ',
        'name': 'meesho'
    },
    'AMAZON': {
        'default': '07AAICA3918J1CV',
        'name': 'amazon'
    },
    'FLIPKART': {
        'default': '07AABCF1429G1Z8',
        'name': 'flipkart'
    }
}
