# BasketTrade - Automated Basket Order Processing System

## Overview

**BasketTrade** is an intelligent Python application designed to automate and optimize the processing of basket orders provided by clients to brokers. The system significantly reduces processing time from ~20 minutes (manual) to 2-3 minutes while minimizing errors through intelligent validation, automatic correction, and data memorization.

## Key Features

✅ **Automated Order File Processing** - Handles multiple file formats (Excel, CSV) with intelligent formatting  
✅ **Intelligent Data Validation** - Validates order information with automatic error correction when possible  
✅ **Bloomberg API Integration** - Real-time ticker verification and data enrichment via Bloomberg API  
✅ **Cross-Machine Communication** - Distributed system allowing Bloomberg lookups from non-licensed computers  
✅ **Client Data Memory System** - Learns from previous orders to improve accuracy and efficiency over time  
✅ **Audit Trail** - Complete logging and audit tracking for compliance  
✅ **Direct ULink Integration** - Seamless export of processed orders to the ULink trading platform  
✅ **Performance** - Reduces processing time by 85% compared to manual handling

## Project Architecture

```
BasketTrade/
├── BT_main.py                          # Entry point - launches the application
├── BasketTrade_ajout.py                # Initialization module (user, file, client setup)
├── Verification.py                     # Thread-based order verification & processing
├── Bloom.py                            # Bloomberg API wrapper
├── watcher_ordibloom_user_a.py         # File watcher for Bloomberg IPC communication
└── data/                               # Client data memory (databases, audit logs)
```

### Architecture Flow

1. **Entry Point** (`BT_main.py`)
   - Launches the application and initializes the system

2. **Initialization** (`BasketTrade_ajout.py`)
   - Loads user credentials and configuration
   - Reads input file (Excel/CSV)
   - Identifies client information

3. **Processing Thread** (`Verification.py`)
   - Formats order data to specification
   - Validates all order information
   - Attempts automatic error correction
   - Suggests solutions for uncorrectable errors
   - Exports to ULink

4. **Bloomberg Integration** (`Bloom.py` + `watcher_ordibloom_user_a.py`)
   - Primary module handles direct API calls when Bloomberg license is available
   - File watcher system enables inter-machine communication:
     - Work computer sends JSON requests to shared folder
     - Bloomberg computer processes requests and returns results
     - Ensures compatibility across different environments

## Technology Stack

| Component | Technology | Purpose |
|-----------|-----------|---------|
| Database | `pysmql` | Client data storage & retrieval |
| Data Processing | `pandas` | Order file parsing and manipulation |
| File Monitoring | `watchdog` | Bloomberg IPC communication |
| Market Data | `blpapi` | Bloomberg Terminal API integration |
| Language | Python 3.x | Core application logic |

## Installation & Setup

### Prerequisites

- Python 3.8+
- Access to Bloomberg Terminal (on licensed machine)
- SQL Server for client database
- ULink API credentials


### Typical Workflow

1. Prepare your basket order file (Excel/CSV format)
2. Place it in the designated input folder
3. Run `BT_main.py`
4. Application will:
   - Detect and load the file
   - Format data to specification
   - Validate all orders against Bloomberg data
   - Auto-correct errors where possible
   - Flag issues for manual review if needed
   - Export to ULink automatically

### Bloomberg Integration (Distributed Setup)

On **Work Computer:**
- Run `BT_main.py` normally
- Application sends validation requests via file watcher

On **Bloomberg Computer:**
- Run `watcher_ordibloom_user_a.py` in the background
- Monitors shared folder for requests
- Executes Bloomberg API calls
- Returns results as JSON files

## Performance Metrics

| Metric | Manual | BasketTrade |
|--------|--------|-------------|
| Processing Time | ~20 minutes | 2-3 minutes |
| Error Rate | Variable | Minimized |
| Time Saved | — | ~85% reduction |

## Current Status

🔶 **Beta** - The application is currently in active development and testing. Core functionality is operational, but additional refinement and edge case handling may occur.


## Database & Memory System

The application maintains client-specific databases containing:
- Historical order patterns
- Common ticker corrections
- Client-specific validation rules
- Transaction audit logs

This data memory allows the system to continuously improve accuracy with each processed order.

## Error Handling

The application handles errors intelligently:

1. **Automatic Corrections** - Common formatting and ticker issues are resolved automatically
2. **Suggested Solutions** - When automatic correction isn't possible, the system provides ranked suggestions
3. **Audit Logging** - All corrections and decisions are logged for compliance and review
4. **User Notification** - Problematic orders are flagged for manual review with detailed explanations

## Technical Highlights

### Multi-Threading Architecture
- Verification processes run in dedicated threads for optimal performance
- Non-blocking file system operations via `watchdog`

### Distributed Computing Model
- Work and Bloomberg machines operate independently
- Asynchronous communication via JSON file exchange
- Fault-tolerant retry mechanisms

### Data Persistence
- SQL-based client memory system
- Scalable to thousands of orders and clients
- Comprehensive audit trail for regulatory compliance

## Contributing

This is a professional trading system. For questions or improvements, please contact the development team.

## License

Proprietary - Internal Use Only

## Contact & Support

For issues, feature requests, or support regarding BasketTrade, please reach out to the development team.

---

**Note:** This application requires proper configuration and credentials to function. Direct Bloomberg license access is required on at least one machine in the deployment. The file watcher system enables seamless integration across different computer environments while maintaining security and compliance standards.
