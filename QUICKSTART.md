# ?? Quick Start Guide - Monchai Insurance TypeScript

## ?? Installation

```bash
cd c:\Project\Monchai
npm install
```

## ?? Running the Application

### Option 1: Development Mode (Recommended)
```bash
npm run dev
```
- ? Hot reload enabled
- ? TypeScript runs directly with `tsx`
- ? Faster development cycle
- ? No need to build

### Option 2: Production Mode
```bash
npm run build    # Compile TypeScript
npm start        # Run compiled JavaScript
```
- ? Optimized for production
- ? Better performance
- ? Runs from `dist/` folder

### Option 3: Original JavaScript (Fallback)
```bash
npm run start:old
```
- ? Uses original `server.js`
- ? No TypeScript needed
- ? Good for emergencies

## ?? Common Commands

```bash
# Development
npm run dev          # Start dev server with hot reload
npm run watch        # Watch TypeScript files for changes

# Production
npm run build        # Compile TypeScript to JavaScript
npm start            # Run compiled version

# Fallback
npm run start:old    # Run original JavaScript version
```

## ?? Where to Find Things

```
src/
??? server.ts                 # ?? START HERE - Main server entry point
??? config.ts                 # ?? Configuration (PORT, URLs, etc.)
??? types/index.ts           # ?? Type definitions
??? utils/
?   ??? helpers.ts          # ??? Utility functions
?   ??? validation.ts       # ? Form validation
??? services/
?   ??? cache.ts            # ?? Performance caching
?   ??? customerService.ts  # ?? Customer operations
?   ??? backupService.ts    # ?? Backup functions
??? handlers/
?   ??? customerHandlers.ts # ?? HTTP request handlers
??? views/
    ??? layout.ts           # ?? HTML layout
    ??? customerViews.ts    # ?? Customer forms
```

## ?? Debugging

### Check if server is running
```bash
# Windows PowerShell
Get-Process | Where-Object {$_.ProcessName -eq "node"}

# Stop all Node processes
Get-Process | Where-Object {$_.ProcessName -eq "node"} | Stop-Process -Force
```

### View logs
```bash
# Terminal shows logs automatically
# Look for:
# ?? TypeScript Server running at http://127.0.0.1:3000
# ?? Caching enabled (30s TTL) for better performance
# ?? Environment: development
```

### Check TypeScript errors
```bash
npm run build    # Will show any TypeScript errors
```

## ?? Common Issues

### Port 3000 already in use
```bash
# Stop existing server
Get-Process | Where-Object {$_.ProcessName -eq "node"} | Stop-Process -Force

# Or change port in .env
PORT=3001
```

### TypeScript compilation errors
```bash
# Check tsconfig.json
# Make sure all imports end with .js (not .ts)
import { something } from './file.js';  # ? Correct
import { something } from './file.ts';  # ? Wrong
```

### Cache issues
```bash
# Clear node_modules and reinstall
rm -rf node_modules
npm install
```

## ?? Performance Monitoring

### Check cache effectiveness
```typescript
// Cache stats are logged in customerService.ts
// Look for console output:
// "Cache HIT" - Data served from cache (fast!)
// "Cache MISS" - Data fetched from API (slower)
```

### Monitor API calls
```bash
# Watch terminal for Google Sheets API calls
# Should see fewer calls after first page load
# Cache TTL: 30 seconds
```

## ?? Development Workflow

### 1. Make Changes
```bash
# Edit files in src/ folder
# NOT in dist/ folder (auto-generated)
```

### 2. Test Changes
```bash
# Dev mode auto-reloads
npm run dev

# Or build and test
npm run build
npm start
```

### 3. Verify Types
```bash
# TypeScript will catch errors
npm run build

# Look for red squiggles in VS Code
```

## ??? Useful VS Code Extensions

1. **TypeScript and JavaScript Language Features** (built-in)
2. **ESLint** - Code linting
3. **Prettier** - Code formatting
4. **Thunder Client** - Test API endpoints

## ?? Environment Variables

Create `.env` file:
```env
PORT=3000
SHEET_WEBHOOK_URL=https://your-apps-script-url
SHEET_DATA_URL=https://your-sheets-data-url
SHEET_UPDATE_URL=https://your-apps-script-url
```

## ?? Security Notes

- ? `.env` file is git-ignored
- ? HTML is escaped to prevent XSS
- ? Request size limited to 1MB
- ? Input validation on all forms

## ?? Performance Tips

### Caching
- Customer data cached for 30s
- Cache automatically invalidates on updates
- Adjust TTL in `src/services/cache.ts`

### API Optimization
- Minimize Google Sheets API calls
- Use batch operations when possible
- Consider implementing request queuing

## ?? Learn More

### TypeScript Basics
```typescript
// Type annotation
const name: string = "John";

// Function with types
function add(a: number, b: number): number {
  return a + b;
}

// Interface
interface Customer {
  name: string;
  phone: string;
}
```

### Project Patterns
- **Services**: Business logic and external API calls
- **Handlers**: HTTP request/response handling
- **Utils**: Reusable helper functions
- **Views**: HTML rendering
- **Types**: TypeScript interfaces

## ?? Getting Help

### Error: "Cannot find module"
```bash
# Check import paths end with .js
# Check file exists in src/
npm run build  # See detailed error
```

### Error: "Type X is not assignable to type Y"
```bash
# Check type definitions in src/types/index.ts
# Use correct types
# TypeScript is strict - this is good!
```

### Server won't start
```bash
# Check .env file exists
# Check PORT is not in use
# Check Node.js version (should be 14+)
node --version
```

## ?? Success Indicators

? Server starts without errors
? Browser opens to http://127.0.0.1:3000
? Home page loads
? Search page shows customers
? Cache working (check terminal logs)
? No TypeScript errors in VS Code

---

**Happy coding!** ??

For questions, check:
- README.md - Full documentation
- MIGRATION_SUMMARY.md - What changed and why
- Source code comments - Inline documentation
