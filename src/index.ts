#!/usr/bin/env node

import { startServer } from './server.js';

const args = process.argv.slice(2);
const useStdio = args.includes('--stdio');
const useHttp = args.includes('--http');

let transportType: 'stdio' | 'http' = 'stdio';

if (useHttp) {
  transportType = 'http';
} else if (!useStdio && !useHttp) {
  // Default to stdio if no argument provided
  transportType = 'stdio';
}

startServer(transportType).catch((error) => {
  console.error('Failed to start server:', error);
  process.exit(1);
});
