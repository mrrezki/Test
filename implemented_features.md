# System Feature Implementation Guide

Below is a detailed, step-by-step documentation of all implemented features across the entire `Task 1` and `Task 2` microservice ecosystem.

---

## 1. Task 1: Core Machine Learning Engine
This microservice serves as the singular authority that actually runs the mathematical algorithms.

* **ML Model Training (`train.py`)**
  * Parses a root dataset (`House Price Dataset.csv`).
  * Transforms the raw target columns utilizing `pandas`.
  * Computes deterministic mathematical predictions securely caching coefficients into a serialized `model.joblib` artifact natively during Docker build times.
* **FastAPI Orchestration (`main.py`)**
  * Mounts an internal lightweight HTTP server across port `8000`.
  * Adheres strictly to **MVC Clean Code** structure routing properties through strict controller interfaces.
  * Implements `pytest` functionality resolving prediction responses cleanly without interacting with the active logic.

---

## 2. Task 2: Sub-System Orchestration (Backend/APIs)
We scaled the infrastructure relying on `Task 1` externally, routing traffic intelligently utilizing two vastly different proxy infrastructures.

### 2.1 Python Property Estimator (App 1)
Serves as an orchestration layer translating validation requests between the portal browser and the local Machine Learning cluster.

* **Pydantic Validation Guardrails (`schemas.py`)**
  * Intercepts incoming network payloads verifying numerical boundaries (i.e. `ensure year_built is <= 2030` and `lot_size is > 0`). Prevents bad characters from breaking native models recursively.
* **Concurrent Communication Hooks (`estimation.py`)**
  * Handles external network timeouts effectively implementing pure Python asynchronous programming (`httpx`).
  * Automatically calculates `Upper Bound` and `Lower Bound` valuation metric properties depending on the exact price calculated by the ML logic container.

### 2.2 Java Market Analytics Server (App 2)
Handles "Heavy Compute" operations, isolating big-dataset functionality out of the Python engine to natively rely upon Java 21's speed execution.

* **Spring Cache Abstractions (`MarketAnalysisService.java`)**
  * Integrates robust `@Cacheable` components locally bypassing SQL overhead entirely for repeated high-traffic requests on global metadata aggregations.
* **Network Export Configurations (`MarketAnalysisController.java`)**
  * Handles raw unstructured properties dumping string-formatted dataset aggregations into literal `text/csv` blob formats.
* **Multi-Stage Container Architecture (`Dockerfile`)**
  * Ensures Maven `target` `.jar` compilations happen invisibly allowing non-Java developers full ability to spin up the container natively.

---

## 3. Task 2: Unified Next.js Portal (Frontend)
Merges all sub-architectures visually together.

* **React Node Internal Networking (`src/lib/api.ts`)**
  * Executes Server Actions masking URL requests natively mapping via standard Environment DNS flags (`TASK1_ML_URL`, `JAVA_API_URL`, etc). This totally negates Cross Origin (CORS) faults!
* **Aesthetic Form Orchestration (`estimations/page.tsx`)**
  * Replaces raw JSON calls traversing the `App 1 (Python)` infrastructure converting the process into rapid numeric entry blocks with CSS disabled state-management rules.
* **Dynamic Dataset Table Layouts (`analysis/page.tsx`)**
  * Consumes raw metadata objects emitted originally by `App 2 (Java)` binding generic objects via `Object.entries(stats)` to automatically draw and filter analytics structures logically.
* **Unified Layout Configurations (`layout.tsx`)**
  * Keeps the structural aesthetics cohesive preserving static header parameters and preventing generic DOM redraws on every internal page route mutation.
