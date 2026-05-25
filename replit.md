# RDO Generator - Engineering Suite

## Overview

The RDO Generator is a web-based application for generating professional Daily Activity Reports (Relatório Diário de Obra - RDO) for construction and engineering projects. The system reads activity data from Google Sheets, processes it through customizable templates, and generates both Excel spreadsheets and Word documents. It features AI-powered text enhancement capabilities to improve activity descriptions and supports batch processing with image attachments.

**Primary Purpose:** Automate the creation of standardized daily construction reports by transforming raw field data into professional, client-ready documentation.

**Tech Stack:** Full-stack TypeScript application using React (Vite), Express.js, and PostgreSQL, with client-side document generation using ExcelJS and Docxtemplater.

## User Preferences

Preferred communication style: Simple, everyday language.

## System Architecture

### Frontend Architecture

**Framework & Tooling:**
- **React 18** with TypeScript via Vite for fast development and optimized builds
- **shadcn/ui components** (New York style) for consistent, accessible UI elements
- **TailwindCSS v4** with custom industrial precision theme (deep slate blue primary, safety yellow secondary)
- **React Query** (@tanstack/react-query) for server state management
- **Wouter** for lightweight client-side routing

**Design Philosophy:**
The application uses a single-page architecture focused on the main dashboard. All document generation happens client-side in the browser, eliminating the need for server-side file processing. This choice prioritizes:
- **User privacy:** No documents stored on server
- **Performance:** Reduces server load and network latency
- **Offline capability:** Core functionality works without constant server connection

**Document Generation Strategy:**
- **ExcelJS:** Generates Excel files with precise cell manipulation, formatting, and image embedding
- **Docxtemplater with PizZip:** Creates Word documents from templates using placeholder substitution
- **JSZip:** Bundles multiple generated files for bulk download
- **File-saver:** Handles client-side file downloads

**Key UI Components:**
- `Header`: Branding and system status
- `FileUpload`: Drag-and-drop interface for template uploads
- `ImageSection`: Manages before/after photo attachments with URL and file support
- `ActivityList`: Displays and allows editing of activity entries with AI enhancement

### Backend Architecture

**Server Framework:**
- **Express.js** on Node.js with TypeScript
- **HTTP-only server** (no WebSocket usage despite ws in dependencies)
- Stateless design - no session management or user authentication required

**API Design:**
The backend serves as a minimal proxy layer with two primary endpoints:

1. **Google Sheets Proxy** (`POST /api/proxy/google-sheet`):
   - **Problem:** Browser CORS restrictions prevent direct fetching of Google Sheets CSV exports
   - **Solution:** Server-side fetch with proper encoding handling (UTF-8)
   - **Benefit:** Maintains client-side data processing while bypassing CORS

2. **AI Rewrite Endpoint** (`POST /api/ai/rewrite`):
   - Enhances activity descriptions using AI services
   - Keeps AI credentials secure on server-side
   - Returns improved text to client for document inclusion

**Static File Serving:**
- Production builds serve pre-compiled React app from `dist/public`
- Development mode uses Vite middleware for HMR and fast refresh

**Build Strategy:**
- esbuild bundles server code to `dist/index.cjs` with selective dependency bundling
- Allowlist strategy for frequently-used dependencies reduces cold start times
- Custom build script (`script/build.ts`) coordinates client and server compilation

### Data Storage Solutions

**No Traditional Database:**
The application is intentionally stateless and database-free. This architectural decision stems from:
- **Use case:** Document generation is a one-time, per-request operation
- **Data source:** All input data comes from external Google Sheets
- **No persistence needs:** Generated files are immediately downloaded
- **Simplicity:** Eliminates database maintenance and migration concerns

**Data Flow:**
1. User provides Google Sheets URL
2. Server proxies CSV data to client
3. Client parses CSV and holds data in React state
4. User customizes and generates documents
5. Files download directly to user's device
6. No data retained after session

**Schema Definition:**
Despite no database, the project uses Drizzle ORM configuration (`drizzle.config.ts`) pointing to PostgreSQL. This suggests the codebase may have been scaffolded from a template or prepared for future database integration. Currently, `shared/schema.ts` only defines Zod validation schemas for API requests, not database schemas.

### Authentication and Authorization

**No Authentication System:**
The application currently has no user authentication, session management, or authorization controls. This is appropriate for:
- Internal tool usage within a trusted network
- Single-user/single-team scenarios
- MVP/prototype phase

**Security Considerations:**
- API endpoints are publicly accessible
- No rate limiting implemented beyond basic Express middleware
- Google Sheets URLs must be public or shared links
- Generated documents may contain sensitive construction data

**Future Enhancement Path:**
The presence of session-related dependencies (`express-session`, `connect-pg-simple`, `passport`, `passport-local`) indicates potential plans for adding authentication.

### External Dependencies

**Third-Party Services:**

1. **Google Sheets API:**
   - **Purpose:** Data source for construction activities
   - **Integration:** CSV export endpoint (no OAuth required for public sheets)
   - **Format:** Expects specific column structure (Data, Técnico Responsável, Atividade Realizada, etc.)

2. **AI Enhancement Service:**
   - **Provider:** Configurable (suggested: Google Generative AI or OpenAI based on dependencies)
   - **Purpose:** Improve activity description clarity and professionalism
   - **Implementation:** Server-side API calls with client-requested text

**Document Processing Libraries:**

1. **ExcelJS:**
   - Advanced Excel file manipulation
   - Supports cell formatting, images, formulas, and styling
   - Chosen over simpler alternatives (like xlsx) for precise template control

2. **Docxtemplater:**
   - Template-based Word document generation
   - Placeholder substitution (e.g., `{{tecnico}}`, `{{data}}`)
   - Maintains document formatting and editability

3. **XLSX (SheetJS):**
   - Reading uploaded Excel templates
   - Lightweight parsing for template validation

**Image Handling:**
- `docxtemplater-image-module-free`: Embeds images in Word documents
- Support for both file uploads and URL-based images
- Proxy endpoint for fetching remote images to avoid CORS

**Development Tools:**
- **Replit-specific plugins:** Cartographer (code navigation), dev banner, runtime error modal
- **Vite plugins:** Custom `metaImagesPlugin` for OpenGraph image URL handling
- **Font loading:** Google Fonts (Barlow, Inter, JetBrains Mono) for consistent typography

**Key Design Decisions:**

- **Client-side generation:** Reduces server complexity, improves privacy
- **Template-based approach:** Users can customize document appearance without code changes
- **Google Sheets integration:** Leverages familiar data entry tool for field teams
- **Modular UI components:** shadcn/ui provides consistency and accessibility out-of-the-box
- **Industrial theme:** Color scheme chosen to reflect construction/engineering domain