import app from './src/app.js';
import { config } from './src/config/env.js';

app.listen(config.port, () => {
  console.log(`====================================================`);
  console.log(`🚀 Docs Service Server running on port ${config.port}`);
  console.log(`🔗 Local URL: http://localhost:${config.port}`);
  console.log(`====================================================`);
});
