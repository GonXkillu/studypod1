import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  AlignmentType,
  Table,
  TableRow,
  TableCell,
  WidthType,
  ShadingType,
  PageBreak,
} from "docx";
import { writeFileSync } from "fs";

// ─── helpers ────────────────────────────────────────────────────────────────

function h1(text) {
  return new Paragraph({
    text,
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 400, after: 120 },
  });
}

function h2(text) {
  return new Paragraph({
    text,
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 300, after: 100 },
  });
}

function h3(text) {
  return new Paragraph({
    text,
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 200, after: 80 },
  });
}

function body(text) {
  return new Paragraph({
    children: [new TextRun({ text, size: 24 })],
    spacing: { before: 100, after: 120, line: 360 },
    alignment: AlignmentType.JUSTIFIED,
  });
}

function blank() {
  return new Paragraph({ text: "" });
}

function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

// ─── table helpers ───────────────────────────────────────────────────────────

function tableHeader(cells) {
  return new TableRow({
    children: cells.map(
      (text) =>
        new TableCell({
          children: [
            new Paragraph({
              children: [
                new TextRun({ text, bold: true, size: 18, color: "FFFFFF" }),
              ],
              alignment: AlignmentType.CENTER,
            }),
          ],
          shading: { type: ShadingType.SOLID, color: "1F3864" },
          margins: { top: 80, bottom: 80, left: 80, right: 80 },
        })
    ),
  });
}

function tableRow(cells, shade = false) {
  return new TableRow({
    children: cells.map(
      (text) =>
        new TableCell({
          children: [
            new Paragraph({
              children: [new TextRun({ text, size: 18 })],
              alignment: AlignmentType.LEFT,
            }),
          ],
          shading: shade
            ? { type: ShadingType.SOLID, color: "DCE6F1" }
            : undefined,
          margins: { top: 60, bottom: 60, left: 80, right: 80 },
        })
    ),
  });
}

// ─── timeline table ──────────────────────────────────────────────────────────

function timeline() {
  const headers = [
    "Phase",
    "Milestone / Task",
    "Resources",
    "Dependencies",
    "Deliverable",
    "Description",
    "Dates w/ Duration",
  ];
  const rows = [
    [
      "Planning",
      "Requirements Definition",
      "Developer",
      "None",
      "Requirements Document",
      "Define all functional and non-functional requirements for the application based on the approved project scope.",
      "05/16/2026 – 05/19/2026 (4 days)",
    ],
    [
      "Planning",
      "Project Scope & Risk Analysis",
      "Developer",
      "Requirements Document",
      "Risk Register",
      "Identify project risks, assess likelihood and impact, and document mitigation strategies.",
      "05/19/2026 – 05/21/2026 (2 days)",
    ],
    [
      "Design",
      "UI/UX Wireframes & Mockups",
      "Developer",
      "Requirements Document",
      "Wireframe Files",
      "Create low-fidelity wireframes and high-fidelity mockups for all application pages.",
      "05/21/2026 – 05/23/2026 (2 days)",
    ],
    [
      "Design",
      "Database Schema Design",
      "Developer",
      "Requirements Document",
      "ER Diagram",
      "Design the PostgreSQL database structure including tables, columns, and relationships between entities.",
      "05/21/2026 – 05/23/2026 (2 days)",
    ],
    [
      "Development – Sprint 1",
      "User Authentication (register/login)",
      "Developer",
      "Wireframes, DB Schema",
      "Auth Module",
      "Build registration, login, logout, and session management for the application.",
      "05/23/2026 – 05/25/2026 (2 days)",
    ],
    [
      "Development – Sprint 2",
      "Room Listing & Availability View",
      "Developer",
      "Auth Module",
      "Room Availability Feature",
      "Build the room listing page that displays all rooms and their available and reserved time slots.",
      "05/25/2026 – 05/27/2026 (2 days)",
    ],
    [
      "Development – Sprint 3",
      "Reservation CRUD & Conflict Checking",
      "Developer",
      "Room Availability Feature",
      "Reservation Module",
      "Build reservation creation, cancellation, and scheduling conflict prevention features.",
      "05/27/2026 – 05/29/2026 (2 days)",
    ],
    [
      "Development – Sprint 4",
      "Admin Dashboard (manage rooms & bookings)",
      "Developer",
      "Reservation Module",
      "Admin Panel",
      "Build the admin dashboard for managing rooms and all student reservations system-wide.",
      "05/29/2026 – 05/31/2026 (2 days)",
    ],
    [
      "Testing",
      "Unit & Integration Testing",
      "Developer",
      "All Dev Sprints Complete",
      "Test Reports",
      "Write and execute unit and integration tests for all critical backend functions and API endpoints.",
      "05/31/2026 – 06/02/2026 (2 days)",
    ],
    [
      "Testing",
      "Developer Acceptance Testing",
      "Developer",
      "Unit/Integration Tests",
      "Acceptance Test Checklist",
      "Systematically verify every user-facing feature against the requirements checklist and resolve defects.",
      "06/02/2026 – 06/03/2026 (1 day)",
    ],
    [
      "Deployment",
      "Containerization with Docker",
      "Developer",
      "Acceptance Test Checklist",
      "Docker Image",
      "Author a Dockerfile and Docker Compose file; build and validate the container image locally.",
      "06/03/2026 – 06/04/2026 (1 day)",
    ],
    [
      "Deployment",
      "Cloud Deployment to Vercel",
      "Developer",
      "Docker Image",
      "Live Production URL",
      "Connect GitHub repository to Vercel, configure environment variables, and deploy the production application.",
      "06/04/2026 – 06/05/2026 (1 day)",
    ],
    [
      "Post-Deployment",
      "Documentation & Final Review",
      "Developer",
      "Deployment Complete",
      "Final Documentation Package",
      "Compile all project documents — requirements, design artifacts, test report, and user guide — into the final submission.",
      "06/05/2026 – 06/07/2026 (2 days)",
    ],
  ];

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      tableHeader(headers),
      ...rows.map((r, i) => tableRow(r, i % 2 === 0)),
    ],
  });
}

// ─── environment / cost table ────────────────────────────────────────────────

function envTable() {
  const headers = ["Tool / Technology", "Purpose", "Cost"];
  const rows = [
    [
      "HTML, CSS, JavaScript",
      "Frontend — pages, styling, and interactivity without a framework",
      "Free / Open Source",
    ],
    [
      "Node.js + Express.js",
      "Backend runtime and web framework for the REST API",
      "Free / Open Source",
    ],
    [
      "PostgreSQL (Neon)",
      "Relational database for storing users, rooms, and reservations",
      "Free tier (Neon managed cloud)",
    ],
    [
      "Docker",
      "Containerization of the application for consistent deployment",
      "Free (Docker Desktop community edition)",
    ],
    [
      "Vercel",
      "Cloud hosting platform for deploying the production application",
      "Free tier (Hobby plan)",
    ],
    [
      "Visual Studio Code",
      "Primary IDE for writing and debugging code",
      "Free / Open Source",
    ],
    [
      "Git + GitHub",
      "Version control and source code repository",
      "Free (public repository)",
    ],
    [
      "Jest / Node:test",
      "Unit and integration testing framework for JavaScript",
      "Free / Open Source",
    ],
    [
      "Laptop / Development Machine",
      "Hardware used for all software development and testing",
      "~$800 (existing hardware, estimated value)",
    ],
  ];

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      tableHeader(headers),
      ...rows.map((r, i) => tableRow(r, i % 2 === 0)),
    ],
  });
}

// ─── document assembly ───────────────────────────────────────────────────────

const doc = new Document({
  styles: {
    default: {
      document: { run: { font: "Calibri", size: 24 } },
    },
    paragraphStyles: [
      {
        id: "Heading1",
        name: "Heading 1",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { size: 32, bold: true, color: "1F3864" },
        paragraph: { spacing: { before: 400, after: 200 } },
      },
      {
        id: "Heading2",
        name: "Heading 2",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { size: 26, bold: true, color: "2E5FA3" },
        paragraph: { spacing: { before: 280, after: 120 } },
      },
      {
        id: "Heading3",
        name: "Heading 3",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { size: 24, bold: true, color: "2E75B6" },
        paragraph: { spacing: { before: 200, after: 80 } },
      },
    ],
  },
  sections: [
    {
      children: [
        // ── Cover Page ──
        blank(),
        blank(),
        blank(),
        new Paragraph({
          children: [
            new TextRun({
              text: "D424 – Software Engineering Capstone",
              bold: true,
              size: 36,
              color: "1F3864",
            }),
          ],
          alignment: AlignmentType.CENTER,
          spacing: { after: 200 },
        }),
        new Paragraph({
          children: [
            new TextRun({ text: "Task 2: Capstone Proposal", size: 30, color: "2E5FA3" }),
          ],
          alignment: AlignmentType.CENTER,
          spacing: { after: 200 },
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: "StudyPod Scheduler",
              bold: true,
              size: 40,
              color: "1F3864",
            }),
          ],
          alignment: AlignmentType.CENTER,
          spacing: { after: 400 },
        }),
        new Paragraph({
          children: [new TextRun({ text: "Student Name: Amal Osman", size: 24 })],
          alignment: AlignmentType.CENTER,
          spacing: { after: 100 },
        }),
        new Paragraph({
          children: [new TextRun({ text: "Student ID: 012291062", size: 24 })],
          alignment: AlignmentType.CENTER,
          spacing: { after: 100 },
        }),
        new Paragraph({
          children: [new TextRun({ text: "Date: May 2026", size: 24 })],
          alignment: AlignmentType.CENTER,
          spacing: { after: 100 },
        }),
        new Paragraph({
          children: [new TextRun({ text: "Instructor: Candice Allen", size: 24 })],
          alignment: AlignmentType.CENTER,
          spacing: { after: 100 },
        }),

        pageBreak(),

        // ══════════════════════════════════════════════════════════════
        // SECTION A
        // ══════════════════════════════════════════════════════════════
        h1("Section A: Business Problem or Opportunity"),

        h2("The Customer"),
        body(
          "The StudyPod Scheduler application is designed for academic libraries and study-space facilities that currently manage room reservations through manual, paper-based processes. The typical user of this system is a university or college student between the ages of 18 and 30 who regularly needs access to a quiet, dedicated study space in a shared library environment. These students are generally comfortable navigating websites and mobile-friendly web applications, but they do not have access to any existing online reservation tool through their institution. In many cases, the library they attend has limited IT resources and no dedicated software team, meaning that any scheduling system currently in place is informal and managed by general staff rather than a specialized administrator."
        ),
        body(
          "This project was inspired by observing study space scheduling practices in Somali-based libraries, where room reservations are still handled manually using paper sign-up sheets maintained at the front desk or through informal verbal arrangements with library staff. During busy study periods — such as the weeks preceding midterm and final examinations — this approach becomes disorganized and difficult to manage, leaving students uncertain about availability and creating unnecessary friction in their academic routines. The application is designed with this demographic and context in mind, though it is general-purpose enough to serve any small-to-medium-sized academic library that faces the same challenge."
        ),
        body(
          "The expected organizational profile for an institution using this application is a library with a small staff, between two and ten rooms or study pods available for reservation, and a student population in the hundreds to low thousands. The institution's mission is to support student academic success by providing quiet, organized study spaces, and its goal is to make those spaces more accessible and easier to manage without requiring significant investment in expensive proprietary software."
        ),

        h2("Business Case"),
        body(
          "The core problem this application addresses is that manual room scheduling systems — whether paper-based or managed through informal verbal agreements — are inherently disorganized and inaccessible to students outside of library operating hours. When a student wants to reserve a study room, they are typically required to travel to the library in person, locate the sign-up sheet or speak with a staff member, and hope that a suitable time slot is still available. If all rooms are fully booked for their desired time, they have no way to learn when a room will next become available without waiting on-site or making repeat visits. During high-demand periods, this process creates congestion, frustration, and wasted time for both students and staff."
        ),
        body(
          "Beyond the inconvenience for students, the absence of a centralized digital system means that libraries have no reliable way to track reservation history, analyze room utilization patterns, or identify scheduling conflicts before they occur. Cancellations are frequently not communicated in a timely manner, resulting in rooms sitting unused while other students are turned away. These inefficiencies represent a meaningful opportunity cost for institutions whose primary purpose is to support student learning."
        ),
        body(
          "The StudyPod Scheduler resolves these problems by providing a centralized web-based platform where students can check availability and make reservations at any time from any internet-enabled device, without needing to visit the library in person. The application includes built-in conflict checking that prevents overlapping bookings, and it provides an administrative interface through which authorized users can manage rooms and reservations directly. Because the application is deployed to the cloud and accessible via a public URL, it requires no special hardware or software installation on the part of the library or its students."
        ),
        body(
          "No comparable free or low-cost solution currently serves this market segment in the way this application is designed to. Commercially available room-booking platforms either require paid licensing, are designed for corporate office environments, or assume a level of IT infrastructure that small academic libraries in developing regions do not have. The StudyPod Scheduler is purpose-built for this use case and is designed to be lightweight, accessible, and straightforward to use for students with varying levels of technical experience."
        ),

        h2("Fulfillment"),
        body(
          "The StudyPod Scheduler will fulfill the needs of its users through a clean, browser-based interface that works on any device with an internet connection, including smartphones, which are the primary computing devices for many students in the target demographic. The application's pages will be built with standard HTML, CSS, and JavaScript, keeping the interface lightweight and fast-loading even on slower internet connections."
        ),
        body(
          "A new user will begin by creating an account using their name, email address, and a password. Returning users will log in through a standard email-and-password form. Once authenticated, a student will land on the room availability page, which displays all study rooms and pods available in the system, their capacity, and a view of currently reserved and open time slots. The student can select an available slot, fill in a short reservation form, and submit it to create a booking. The application will display a confirmation immediately after a successful reservation. Students can also view all of their active bookings from a personal reservations page and cancel any upcoming reservation with a single action."
        ),
        body(
          "The system enforces reservation integrity by checking for scheduling conflicts at the moment a booking is submitted. If the requested time slot is already taken for the selected room, the server will reject the request and return an error message, prompting the user to choose a different time. This logic is enforced at the backend API level and validated against the database, ensuring that simultaneous submission attempts cannot result in double bookings."
        ),
        body(
          "An administrative section of the application — protected by a role-based access control system — provides a dashboard through which an administrator can add, edit, or remove study rooms; view all reservations across all rooms; and cancel or modify any existing booking. The frontend communicates with the backend through a RESTful HTTP API, and all data is stored in a PostgreSQL relational database. All communication between the client and server occurs over HTTPS to ensure that user credentials and reservation data are transmitted securely."
        ),

        pageBreak(),

        // ══════════════════════════════════════════════════════════════
        // SECTION B
        // ══════════════════════════════════════════════════════════════
        h1("Section B: Software Development Life Cycle Methodology"),

        h2("SDLC Methodology Overview"),
        body(
          "The StudyPod Scheduler project will be developed using the Agile software development life cycle methodology. Agile is an iterative, incremental approach to software development that emphasizes adaptability, continuous progress, and the delivery of working software in short development cycles called sprints (Sommerville, 2016). Rather than attempting to fully specify and build the entire system before delivering anything testable, Agile breaks the project into a series of focused sprints, each producing a functional increment of the application that can be evaluated before the next sprint begins."
        ),
        body(
          "Agile is well suited to this project for two primary reasons. First, while the high-level requirements for the StudyPod Scheduler are clearly defined, the specific behavior of individual features — such as the reservation conflict-checking logic and the layout of the administrative dashboard — may require iteration and adjustment as development proceeds and the developer reviews each working increment. Second, because this is a solo development project, the lightweight and self-directed nature of Agile is more practical than a formal methodology that assumes multiple team members and extended documentation ceremonies. The developer will serve as the sole member of the development team as well as the project lead, responsible for planning, executing, and reviewing all sprint work."
        ),
        body(
          "The project is organized into five phases: Planning, Design, Development, Testing, and Deployment. Each phase produces specific deliverables that build upon the previous phase's outputs, creating a traceable progression from the initial requirements to a production-ready, cloud-hosted application."
        ),

        h2("Phase Descriptions"),

        h3("Phase 1: Planning"),
        body(
          "During the Planning phase, the developer will define the full functional and non-functional requirements for the application based on the approved project scope from Task 1. This involves documenting the exact features the system must support, identifying any constraints or assumptions, and producing a risk register that captures potential challenges — such as unexpected complexity in the conflict-checking logic or deployment configuration issues — along with proposed mitigation strategies. The requirements document produced in this phase is the foundation for all subsequent design and development decisions and must be completed before any design work begins."
        ),

        h3("Phase 2: Design"),
        body(
          "In the Design phase, the developer will translate the requirements document into concrete technical specifications. This includes creating wireframes and a high-fidelity mockup of each page in the application, designing the PostgreSQL database schema with entity-relationship diagrams, and defining the RESTful API contract that specifies the endpoints, HTTP methods, and expected request and response formats. The Design phase depends on the completed requirements document and must itself be completed before any code is written, as both the database schema and the API design serve as direct references during development."
        ),

        h3("Phase 3: Development"),
        body(
          "The Development phase is organized into four iterative sprints. Sprint 1 delivers the user authentication module, covering account registration, login, logout, and session management. Sprint 2 delivers the room listing and availability view that allows authenticated users to browse available rooms and time slots. Sprint 3 delivers the reservation creation, cancellation, and scheduling conflict prevention features. Sprint 4 delivers the administrative dashboard, which allows an authorized administrator to manage rooms and all reservations in the system. Each sprint depends on the successful completion of the sprint before it, as later features rely on the foundation established by earlier ones."
        ),

        h3("Phase 4: Testing"),
        body(
          "The Testing phase begins after all four development sprints are complete. The developer will write and execute unit tests for all critical backend functions and integration tests that verify the complete request-to-database-to-response cycle for each major API endpoint. Following automated testing, the developer will perform a structured developer acceptance test by working through a prepared checklist of every user-facing feature defined in the requirements document and verifying that each one behaves as specified. The Testing phase has a hard dependency on the completion of all development sprints, and no deployment activities may begin until all critical test findings have been addressed."
        ),

        h3("Phase 5: Deployment"),
        body(
          "In the Deployment phase, the completed and tested application will be containerized using Docker and deployed to the Vercel cloud hosting platform. This phase also includes the preparation of all final project documentation. The Deployment phase depends on the successful completion of the Testing phase and the resolution of any critical defects identified during testing."
        ),

        h2("B1: Deliverables"),

        h3("Project Deliverables"),
        body(
          "The requirements document is the primary project deliverable of the Planning phase. It is authored by the developer and serves as the authoritative reference for all design and development decisions. The document captures all functional requirements — such as user authentication, room availability display, reservation creation, conflict checking, and administrative management — as well as non-functional requirements such as security, response time, and browser compatibility. This document is a dependency for all subsequent phases."
        ),
        body(
          "The risk register is a project deliverable produced during the Planning phase. It identifies potential risks to the project timeline and technical quality, estimates the likelihood and impact of each, and documents a mitigation strategy for each identified risk. It will be reviewed at the start of each phase to determine whether any risks have materialized and whether any mitigation actions need to be applied."
        ),
        body(
          "The test report is the primary project deliverable of the Testing phase. It documents the results of all unit tests, integration tests, and the developer acceptance testing checklist, including the outcome of each test case and the resolution of any defects identified. The test report is a dependency for the Deployment phase, as it serves as the quality gate that must be passed before the production deployment is authorized."
        ),
        body(
          "The final documentation package is produced during the Deployment phase. It consolidates the requirements document, risk register, design artifacts, test report, technical architecture overview, and deployment steps into a single organized submission that satisfies the capstone documentation requirements."
        ),

        h3("Product Deliverables"),
        body(
          "The wireframe and UI mockup files are product deliverables of the Design phase. They provide the visual blueprint for the application's pages — including the login screen, the room availability dashboard, the reservation form, the personal bookings page, and the admin panel — and serve as the direct reference for building the HTML and CSS during development."
        ),
        body(
          "The database entity-relationship diagram is a product deliverable of the Design phase. It defines the structure of the PostgreSQL database, including the tables for users, rooms, and reservations, their columns and data types, and the foreign key relationships between them. The database schema is created directly from this diagram during Sprint 1."
        ),
        body(
          "The authentication module is a product deliverable of Sprint 1. It encompasses the Express.js API routes for registration and login, the password hashing implementation using bcrypt for secure credential storage, the session management middleware, and the HTML/CSS/JavaScript frontend forms for the login and registration pages."
        ),
        body(
          "The reservation module is a product deliverable of Sprint 3. It includes the backend API routes for creating, retrieving, and canceling reservations, the scheduling conflict-checking logic that queries the database to prevent overlapping bookings, and the frontend forms and pages that allow users to interact with their reservations."
        ),
        body(
          "The Docker image is a product deliverable of the Deployment phase. It is a containerized package of the complete application — including the Node.js backend server, the compiled static frontend files, and all environment configuration — that can be deployed consistently to the Vercel cloud hosting environment."
        ),
        body(
          "The production deployment is the final product deliverable. It is the live, publicly accessible version of the StudyPod Scheduler application hosted on Vercel, connected to the production PostgreSQL database instance, and served over HTTPS."
        ),

        h2("B2: Deployment Plan and Outcomes"),

        h3("Application Description"),
        body(
          "The StudyPod Scheduler is a full-stack web application built for academic libraries and study-space facilities that need a simple, centralized system for managing room reservations. The application follows a client-server architecture: the frontend is built with plain HTML, CSS, and JavaScript and is served as static files, while the backend is a RESTful API built with Node.js and Express.js that handles all business logic, authentication, and data access. Reservation and user data are stored in a PostgreSQL relational database hosted through Neon. The application is containerized using Docker and deployed to the Vercel cloud hosting platform, where it is accessible to users through any modern web browser at a public HTTPS URL. Once deployed, students use the application to view available study rooms, create and cancel reservations, and manage their bookings online. Administrators use a role-protected dashboard within the same application to manage rooms, view all reservations, and make modifications as needed. No software installation is required on the user's device."
        ),

        h3("Production Environment and Hosting Setup"),
        body(
          "The deployment of the StudyPod Scheduler to the production environment will take place at the conclusion of the Testing phase, after the developer has confirmed through the acceptance testing checklist that all features specified in the requirements document function as intended. This sequencing ensures that the version of the application deployed to production has been systematically verified against every defined requirement before it is made publicly accessible."
        ),
        body(
          "The deployment process begins with containerization. The developer will author a Dockerfile that specifies the Node.js runtime version, installs all application dependencies defined in the project's package.json file, copies the application source files into the container, and defines the command that starts the Express backend server. A Docker Compose file will also be created to define the local multi-container environment that includes both the application server and a local PostgreSQL instance, which will be used for pre-deployment testing. Once the Docker image is confirmed to work correctly in the local environment, the developer will push it to a container registry."
        ),

        h3("Database Deployment and Security Configuration"),
        body(
          "For the Vercel deployment, the developer will configure the project through the Vercel dashboard, connecting the GitHub repository so that each push to the main branch triggers an automatic deployment pipeline. All sensitive configuration values — including the PostgreSQL database connection string and the session secret key — will be stored as environment variables within the Vercel project settings rather than hardcoded in the source code, in accordance with security best practices. The production database will be a cloud-managed PostgreSQL instance provided through Neon, which integrates directly with Vercel and provides automated daily backups and point-in-time recovery, ensuring that reservation data is protected against accidental loss. All traffic between users and the application is encrypted using HTTPS, which Vercel provisions and renews automatically through SSL certificates. Role-based access control within the application ensures that administrative routes are accessible only to accounts with the administrator role, preventing unauthorized access to room management and reservation oversight features."
        ),

        h3("Post-Deployment Testing and Validation"),
        body(
          "Before the application is considered fully deployed, the developer will perform a post-deployment smoke test by navigating to the live Vercel URL and manually confirming that each major feature is operational in the production environment: user registration, login, room availability display, reservation creation, conflict detection, reservation cancellation, and administrative dashboard access. This smoke test represents the final verification step before the deployment is considered complete. Any issue discovered during the smoke test will be diagnosed by tracing the failure through the application logs available in the Vercel dashboard, identifying the root cause, and applying a corrective code change before the deployment is finalized. The smoke test has a dependency on the completion of both the Docker containerization and the Vercel configuration steps."
        ),

        h3("Monitoring, Backups, and Recovery"),
        body(
          "After the application is live, the developer will use Vercel's built-in deployment logs and runtime function logs to monitor for any server-side errors or unexpected behavior in production. If a deployment introduces a critical bug, Vercel's instant rollback feature allows the developer to revert the production environment to the previous stable deployment within minutes, minimizing downtime. The Neon PostgreSQL service provides automated daily backups of the production database, and the Neon dashboard allows the developer to restore the database to any point within the retention window in the event of data loss or corruption. These measures ensure that the application can recover quickly from both application-level failures and data-level incidents."
        ),

        h3("Supporting Documentation"),
        body(
          "Two supporting documents will be prepared as part of the final documentation package and will be available to any user or administrator who accesses the application after deployment. The administrator user guide is a written document that explains how to log in to the administrative dashboard, add and edit rooms, view and manage all reservations, and cancel bookings on behalf of students. It is intended for any person responsible for managing the library's study space schedule and will be stored as a PDF accessible from the project's GitHub repository. The technical architecture document describes the application's component structure, database schema, API endpoints, environment variable configuration, and deployment pipeline. It is intended for any developer who needs to maintain, update, or extend the application after the initial deployment and will also be stored in the project's GitHub repository alongside the source code."
        ),

        h3("Deployment Outcome"),
        body(
          "The outcome of deployment is a fully operational, publicly accessible web application available to students and administrators through a browser at any time from any internet-connected device, with no installation required. Students will be able to register, log in, view room availability, create and cancel reservations, and manage their bookings entirely online, replacing the need to visit the library in person to check a paper schedule. Administrators will be able to manage rooms and all reservations through a protected dashboard. The outcome of the supporting documentation is that both students and administrators have clear reference materials to guide their use of the system, and the developer has a complete technical record of the application's architecture that supports future maintenance and updates."
        ),

        h2("B3: Projected Timeline"),
        body(
          "The following table presents the projected timeline for the StudyPod Scheduler project, organized by phase and milestone. All dates reflect the anticipated schedule based on the approved project completion date. Resources assigned to each activity refer to the developer, who is the sole team member responsible for all phases of this project. Dependencies between tasks are noted to indicate which prior deliverables must be completed before a given task can begin."
        ),
        blank(),
        timeline(),
        blank(),
        body(
          "The overall project is estimated to span approximately three weeks from initial planning through final documentation, with a total projected effort of 75 hours distributed across planning (15 hours), development (40 hours), and documentation (20 hours). The most critical dependency in the timeline is the completion of all four development sprints before testing begins, as both unit testing and developer acceptance testing require a fully integrated, functional application."
        ),

        pageBreak(),

        // ══════════════════════════════════════════════════════════════
        // SECTION C
        // ══════════════════════════════════════════════════════════════
        h1("Section C: Programming Environments"),

        h2("Programming Environment"),
        body(
          "The StudyPod Scheduler application will be developed using a straightforward, lightweight technology stack that prioritizes simplicity and accessibility. This stack was deliberately chosen to minimize complexity while still satisfying all full-stack software engineering requirements: a browser-based frontend, a server-side backend, a relational database, containerization, and cloud deployment."
        ),
        body(
          "The frontend of the application will be built using plain HTML, CSS, and JavaScript without a frontend framework. This choice was made deliberately to keep the frontend as simple as possible, since the application's user interface consists of a small number of well-defined pages — the login and registration pages, the room availability dashboard, the reservation form, the personal bookings page, and the admin panel — none of which require the complexity of a reactive component framework. Vanilla JavaScript will be used to make asynchronous HTTP requests to the backend API and to dynamically update the page content based on the server's responses. This approach results in a fast-loading, lightweight frontend that works correctly in any modern browser and does not require a build step."
        ),
        body(
          "The backend of the application will be built using Node.js as the server-side JavaScript runtime and Express.js as the web application framework. Node.js is chosen because it allows the developer to write both the frontend and backend of the application in the same language — JavaScript — which reduces context switching and simplifies the overall development workflow (OpenJS Foundation, 2023). Express.js is a minimal and unopinionated framework for Node.js that provides a clean API for defining HTTP routes and middleware, making it well suited to the RESTful API architecture that the application requires. The backend will handle user authentication, enforce reservation conflict checking, and serve as the single point of communication between the frontend and the database."
        ),
        body(
          "The application's data will be stored in a PostgreSQL relational database. PostgreSQL is a powerful, open-source relational database system with strong support for transactions and data integrity constraints (The PostgreSQL Global Development Group, 2023). Its ACID-compliant transaction support is particularly important for the reservation feature, where concurrent booking attempts must be handled correctly to prevent two users from simultaneously reserving the same room for the same time slot. In the production environment, PostgreSQL will be hosted through Neon, a cloud-managed PostgreSQL service that integrates directly with Vercel and provides a generous free tier suitable for this application's expected usage."
        ),
        body(
          "The application will be containerized using Docker. Docker packages the Node.js application and all of its dependencies into a portable, self-contained image that behaves consistently regardless of the environment in which it is run (Merkel, 2014). This means that the application will behave identically in the developer's local environment and in the Vercel production environment, eliminating a common source of deployment failures. Docker Compose will be used to manage the local development environment, defining both the application container and a local PostgreSQL container as a coordinated service."
        ),
        body(
          "The production application will be deployed to Vercel, a cloud hosting platform that natively supports Node.js applications and provides automatic HTTPS, global content delivery, and continuous deployment triggered by pushes to the project's GitHub repository. Visual Studio Code will serve as the primary integrated development environment throughout the project, and Git with GitHub will be used for version control to maintain a complete history of all changes made during development."
        ),

        h2("Environment Costs"),
        body(
          "The following table summarizes all tools, technologies, and hardware required for the development and deployment of the StudyPod Scheduler application, along with their associated costs. The majority of the tools used in this project are free and open source, making the overall development cost consist almost entirely of developer labor."
        ),
        blank(),
        envTable(),
        blank(),
        body(
          "The total estimated startup infrastructure cost for this project is approximately $800, which represents the estimated market value of the developer's existing laptop hardware. All software tools, development frameworks, hosting services, and database services are available at no monetary cost under their respective free tiers or open-source licenses. The only scenario in which ongoing costs would arise is if the application's usage grew to exceed the free tier limits of the Vercel Hobby plan, in which case the Vercel Pro plan would be required at a cost of approximately $20 per month. No proprietary software licenses or paid third-party API services are required for this project."
        ),

        h2("Human Resource Requirements"),
        body(
          "This project is developed entirely by a single developer who serves in all roles: project lead, frontend developer, backend developer, database designer, tester, and deployment engineer. There are no additional team members, contractors, or collaborators involved in the development of this application."
        ),
        body(
          "The developer's total estimated labor is 75 hours across all project phases, as outlined in the project timeline. Using a market rate of $65 per hour for a mid-level full-stack developer working on a freelance contract basis, the total estimated developer labor cost is $4,875 (75 hours × $65 per hour). This figure is provided as a budgetary estimate to reflect the realistic cost of developing an application of this scope and is not a cost that will be incurred in practice, since the developer is the student completing this capstone project. There are no other human resource costs associated with this project."
        ),

        pageBreak(),

        // ══════════════════════════════════════════════════════════════
        // SECTION D
        // ══════════════════════════════════════════════════════════════
        h1("Section D: Testing Plan"),

        h2("D1: Methods for Validation and Verification"),
        body(
          "The testing strategy for the StudyPod Scheduler application is built around a clear distinction between verification and validation. The verification will confirm that the software was built correctly — meaning each component functions according to its technical specification. Validation will confirm that the software meets the actual needs it was designed to address — meaning the complete application fulfills the requirements defined at the beginning of the project (IEEE, 1990). For the purposes of this fictional business scenario, the project team consists of a developer who writes and executes all automated tests, a QA tester who reviews test outcomes and manages the defect log, and a product owner who provides final sign-off on acceptance testing results."
        ),

        h3("Verification Methods"),
        body(
          "Unit testing will be the primary verification method for the backend of the application. The developer will write unit tests using Node.js's built-in test runner for all critical functions in the system, with a strong focus on the reservation conflict-checking logic, user authentication, and the input validation functions that verify incoming API request data. Each unit test will isolate a single function and execute it against a set of predefined inputs to verify that the function returns the expected output or throws the expected error. Both valid inputs that should succeed and invalid inputs that should be rejected will be covered, ensuring that functions handle edge cases and error conditions correctly. Unit testing is important because it catches defects at the smallest possible scope, making the root cause of failures easier to identify and fix before errors compound across the codebase."
        ),
        body(
          "Integration testing will verify that the separate layers of the application work correctly together as a system. These tests simulate HTTP requests to each of the backend API endpoints and confirm that the complete request pipeline — from the incoming request, through the route handler, through the database query, and finally to the response — all behave as expected. For example, an integration test for creating a reservation will start with submitting a valid POST request with a room ID and time slot, then query the database directly to confirm that the reservation record was correctly created. A second test for the same endpoint will submit a request for a time slot that is already reserved and confirm that the server responds with an HTTP 409 Conflict status code rather than creating a duplicate booking. Integration tests are run against a dedicated test database instance seeded with known data before each run, ensuring that results are repeatable and isolated from any production data. Integration testing is essential because it validates that individual modules that pass unit tests in isolation still behave correctly when they interact with each other."
        ),
        body(
          "Code review will also serve as a verification method. Before moving from one development sprint to the next, the developer will review the code produced in the completed sprint against a checklist with proper error handling, correct HTTP status codes in all API responses, and adherence to the RESTful API contract defined during the Design phase. This ensures that quality issues are identified and corrected within the sprint rather than accumulating and compounding across the codebase."
        ),

        h3("User Acceptance Testing (UAT)"),
        body(
          "User Acceptance Testing will take place after all development sprints are complete and all automated unit and integration tests are passing. In this fictional scenario, the QA tester will coordinate a structured UAT session with a representative group of student end-users and one library administrator, who will interact with the fully integrated application in the Vercel staging environment. The QA tester is the team member who works directly with users during this phase, providing instructions and observing each session without intervening so that genuine usability issues can be identified. The product owner monitors overall progress and is responsible for reviewing outcomes and determining which findings require action before the final release."
        ),
        body(
          "During UAT, participants will be asked to complete a series of tasks derived directly from the functional requirements: registering a new account, logging in, browsing room availability, creating a reservation, attempting to book an already-reserved slot to verify conflict prevention, canceling a reservation, and — for the administrator participant — adding a new room through the admin dashboard. Each task maps to a specific requirement in the requirements document, ensuring that user requirements are tested and confirmed systematically rather than informally. At the end of the session, participants complete a short feedback form rating ease of use and overall satisfaction. The product owner reviews all feedback, categorizes findings by severity, and determines what must be resolved before deployment and what can be deferred to a future update. UAT is included because it validates the application from the perspective of actual end users, confirming that the system not only functions technically but also meets the practical expectations of the people who will use it."
        ),

        h2("D2: Analysis of Test Results"),
        body(
          "The results of all testing activities will be recorded and analyzed and will serve as a guide for any necessary revisions before the application is deployed. Each category of testing follows a specific process for capturing, reviewing, and acting on its findings."
        ),
        body(
          "For unit and integration tests, the developer will review the output report generated after each test run. This report will document the total number of tests executed, the number that passed, and the number that failed, along with the specific assertion that caused each failure. Any failing test will be investigated before the application advances to the next phase. When analyzing a test failure, the developer examines the error output, traces the execution path through the relevant code, and performs a Root Cause Analysis (RCA) to categorize the source of the failure. Root cause categories include logic errors in the business rules — for example, an off-by-one error in the time overlap calculation — incorrect database queries, missing or incorrect error handling in a route handler, or type mismatches introduced by recent code changes. Performing an RCA before applying a fix is important because jumping directly from a failed test to a code change without understanding the underlying cause risks introducing a patch that resolves the symptom without fixing the actual problem, which can cause the same defect to resurface later."
        ),
        body(
          "Once the root cause has been identified, the developer implements a targeted fix and commits it to a feature branch in the project's Git repository. The QA tester then reviews the change before it is merged into the main branch, confirming that the fix addresses the identified root cause and does not introduce unintended side effects. The relevant unit tests, integration tests, and any related acceptance checklist items are then re-executed against the updated code in the Vercel staging environment — a separate deployment connected to a test database, isolated from the production environment — to confirm that the fix is effective before it is promoted to production. All defects are tracked as issues in the project's GitHub Issues board, where each entry records the defect description, the steps to reproduce it, the severity rating, the assigned responsible party, and the current resolution status. The QA tester is responsible for opening new issues when defects are found and for closing them after a fix has been verified in the staging environment."
        ),
        body(
          "For user acceptance testing, the results of each task will be recorded in a structured acceptance test log that notes the feature being tested, the expected behavior as defined in the requirements document, the actual observed behavior, and the pass or fail outcome. Any item that fails will be logged as a defect in the GitHub Issues board with a severity rating. Critical defects — those that prevent a core feature like reservation creation or user authentication from functioning at all — must be resolved before deployment. Major defects that significantly degrade the user experience should be resolved before deployment if time permits. Minor defects covering cosmetic issues or edge-case behaviors may be addressed in a post-launch revision. If a participant identifies a bug during UAT, the QA tester logs the issue and the developer performs an RCA to determine whether the failure stems from a technical error in the backend logic, a flaw in the frontend interface, or a misunderstanding of the original requirements. This distinction matters because the appropriate fix is different in each case: a backend logic error requires a code change, a UI problem may require a design adjustment, and a requirements misunderstanding may require a discussion with the product owner before any change is made."
        ),
        body(
          "After all identified defects have been resolved and verified in the staging environment, the final test report is compiled by the QA tester. This report merges the results of all testing activities into a single document containing a summary of all tests executed by category, a complete log of all defects identified and their resolution status, and a certification statement confirming that all critical defects have been resolved and that the application meets every requirement in the requirements document. The product owner reviews the final test report and provides formal sign-off, which serves as the exit criterion for the Testing phase and the trigger for the Deployment phase to begin. Testing is considered complete only when the final test report has been reviewed and approved by the product owner, all critical and major defects are resolved and verified in staging, and the automated test suite reports a full pass with no failures. This test report will be included in the final documentation package."
        ),

        pageBreak(),

        // ══════════════════════════════════════════════════════════════
        // SECTION E
        // ══════════════════════════════════════════════════════════════
        h1("Section E: References"),

        new Paragraph({
          children: [
            new TextRun({ text: "Cohn, M. (2004). ", size: 24 }),
            new TextRun({
              text: "User stories applied: For agile software development. ",
              size: 24,
              italics: true,
            }),
            new TextRun({ text: "Addison-Wesley Professional.", size: 24 }),
          ],
          spacing: { before: 100, after: 120 },
          indent: { left: 720, hanging: 720 },
        }),
        new Paragraph({
          children: [
            new TextRun({ text: "IEEE. (1990). ", size: 24 }),
            new TextRun({
              text: "IEEE standard glossary of software engineering terminology ",
              size: 24,
              italics: true,
            }),
            new TextRun({
              text: "(IEEE Std 610.12-1990). IEEE. https://doi.org/10.1109/IEEESTD.1990.101064",
              size: 24,
            }),
          ],
          spacing: { before: 100, after: 120 },
          indent: { left: 720, hanging: 720 },
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: "Merkel, D. (2014). Docker: Lightweight Linux containers for consistent development and deployment. ",
              size: 24,
            }),
            new TextRun({
              text: "Linux Journal, 2014",
              size: 24,
              italics: true,
            }),
            new TextRun({
              text: "(239), 2. https://www.linuxjournal.com/content/docker-lightweight-linux-containers-consistent-development-and-deployment",
              size: 24,
            }),
          ],
          spacing: { before: 100, after: 120 },
          indent: { left: 720, hanging: 720 },
        }),
        new Paragraph({
          children: [
            new TextRun({ text: "OpenJS Foundation. (2023). ", size: 24 }),
            new TextRun({
              text: "Node.js documentation. ",
              size: 24,
              italics: true,
            }),
            new TextRun({
              text: "OpenJS Foundation. https://nodejs.org/en/docs",
              size: 24,
            }),
          ],
          spacing: { before: 100, after: 120 },
          indent: { left: 720, hanging: 720 },
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: "The PostgreSQL Global Development Group. (2023). ",
              size: 24,
            }),
            new TextRun({
              text: "PostgreSQL documentation (version 16). ",
              size: 24,
              italics: true,
            }),
            new TextRun({
              text: "https://www.postgresql.org/docs/",
              size: 24,
            }),
          ],
          spacing: { before: 100, after: 120 },
          indent: { left: 720, hanging: 720 },
        }),
        new Paragraph({
          children: [
            new TextRun({ text: "Sommerville, I. (2016). ", size: 24 }),
            new TextRun({
              text: "Software engineering ",
              size: 24,
              italics: true,
            }),
            new TextRun({ text: "(10th ed.). Pearson Education.", size: 24 }),
          ],
          spacing: { before: 100, after: 120 },
          indent: { left: 720, hanging: 720 },
        }),
      ],
    },
  ],
});

const buffer = await Packer.toBuffer(doc);
writeFileSync(
  "/home/runner/workspace/StudyPod_Scheduler_Task2_Proposal.docx",
  buffer
);
console.log("Document written successfully.");
