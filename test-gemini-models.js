const axios = require('axios');

const test = async () => {
  const models = ['gemini-flash-latest', 'gemini-1.5-flash', 'gemini-pro', 'gemini-2.0-flash'];
  
  for (const model of models) {
    try {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=AIzaSyAyebU67zmND0wwt5ws6ozVv-qDSMdRCmM`;
      console.log('\nTesting model:', model);
      
      const response = await axios.post(url, {
        contents: [{
          parts: [{
            text: 'test'
          }]
        }]
      }, { timeout: 30000 });
      
      console.log('✓ Status:', response.status);
      console.log('✓ Model works!');
      break;
    } catch (err) {
      console.log('✗ Error:', err.response?.status || err.code || err.message);
    }
  }
};

test();
