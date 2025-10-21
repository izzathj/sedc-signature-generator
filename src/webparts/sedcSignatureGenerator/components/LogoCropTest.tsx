import * as React from 'react';

const LogoCropTest: React.FC = () => {
  return (
    <div style={{ padding: '20px', fontFamily: 'Arial, sans-serif' }}>
      <h1>Logo Crop Test</h1>
      
      <h2>Test 1: Basic Crop with Red Border</h2>
      <div style={{ 
        width: '140px', 
        height: '100px', 
        border: '3px solid red', 
        overflow: 'hidden', 
        position: 'relative',
        marginBottom: '30px'
      }}>
        <img 
          src="https://sedc.com.my/wp-content/uploads/2025/08/SEDC-new-logo-2025-scaled.png" 
          alt="SEDC Logo" 
          style={{ 
            width: '140px', 
            height: 'auto',
            position: 'absolute',
            top: '50%',
            transform: 'translateY(-50%)'
          }} 
        />
      </div>
      <p>☝️ Logo should be cropped (top and bottom cut off) inside red box</p>

      <h2>Test 2: Two Column Table</h2>
      <table style={{ 
        borderCollapse: 'collapse',
        border: '1px solid #000',
        borderSpacing: 0,
        width: '600px',
        height: '100px',
        fontFamily: 'Arial, sans-serif'
      }}>
        <tbody>
          <tr>
            {/* Left Column - Logo */}
            <td style={{ 
              width: '160px', 
              padding: '10px',
              borderRight: '3px solid #00a651',
              verticalAlign: 'middle'
            }}>
              <div style={{
                width: '140px',
                height: '100px',
                border: '2px solid red',
                overflow: 'hidden',
                position: 'relative'
              }}>
                <img 
                  src="https://sedc.com.my/wp-content/uploads/2025/08/SEDC-new-logo-2025-scaled.png"
                  alt="SEDC Logo"
                  style={{
                    width: '140px',
                    height: 'auto',
                    position: 'absolute',
                    top: '50%',
                    left: '0',
                    transform: 'translateY(-50%)'
                  }}
                />
              </div>
            </td>
            
            {/* Right Column - Text */}
            <td style={{ 
              padding: '10px',
              verticalAlign: 'top'
            }}>
              <div style={{ fontSize: '11pt', fontWeight: 'bold', color: '#00a651' }}>John Doe</div>
              <div style={{ fontSize: '10pt', color: '#666666' }}>Manager</div>
              <div style={{ fontSize: '10pt', color: '#999999', fontStyle: 'italic' }}>IT Department</div>
              <div style={{ fontSize: '9pt', marginTop: '10px' }}>📧 john@sedc.my</div>
              <div style={{ fontSize: '9pt' }}>📞 082551555</div>
              <div style={{ fontSize: '9pt' }}>📱 +60 12-345 6789</div>
              <div style={{ fontSize: '9pt' }}>📍 Menara SEDC</div>
            </td>
          </tr>
        </tbody>
      </table>
      <p>☝️ Logo should be cropped and vertically centered next to text</p>

      <h2>Test 3: Without Position Absolute</h2>
      <div style={{ 
        width: '140px', 
        height: '100px', 
        border: '3px solid blue', 
        overflow: 'hidden',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        marginBottom: '30px'
      }}>
        <img 
          src="https://sedc.com.my/wp-content/uploads/2025/08/SEDC-new-logo-2025-scaled.png" 
          alt="SEDC Logo" 
          style={{ 
            width: '140px', 
            height: 'auto'
          }} 
        />
      </div>
      <p>☝️ Using flexbox centering instead of absolute positioning</p>

      <h2>Test 4: Object-fit Cover</h2>
      <div style={{ 
        width: '140px', 
        height: '100px', 
        border: '3px solid green', 
        overflow: 'hidden',
        marginBottom: '30px'
      }}>
        <img 
          src="https://sedc.com.my/wp-content/uploads/2025/08/SEDC-new-logo-2025-scaled.png" 
          alt="SEDC Logo" 
          style={{ 
            width: '140px', 
            height: '100px',
            objectFit: 'cover',
            objectPosition: 'center'
          }} 
        />
      </div>
      <p>☝️ Using object-fit: cover (might distort logo)</p>
    </div>
  );
};

export default LogoCropTest;