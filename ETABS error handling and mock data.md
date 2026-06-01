
# System Stability & Cross-Platform Development Architecture 🛡️

The ETABS COM API is powerful but highly susceptible to environmental interruptions (e.g., ghost processes, hidden dialog boxes, locked files). Furthermore, structural comparisons introduce high risks of topology mismatches (elements being deleted or added between iterations).

This document outlines the strict error-handling protocols and the cross-platform mock data stub required to ensure the application remains stable and testable in any environment.

---

## 🛑 Part 1: Comprehensive Error Handling

Error handling in this application is broken down into four distinct defensive layers to prevent silent failures and infinite loading states.

### 1. Pre-Flight File Validation

Before the COM API is even initialized, the application must verify write access to the target output files. This prevents fatal `PermissionError` crashes at the very end of the extraction process.

```python
import os

def check_word_file_lock(filepath):
    """Verifies write access to the target Word file before extraction."""
    if os.path.exists(filepath):
        try:
            # Attempt to rename the file to itself to test lock status
            os.rename(filepath, filepath)
            return True
        except OSError:
            return False
    return True
```


### 2. COM API Process Sweeping

ETABS crashes frequently leave "ghost"** **`ETABS.exe` processes running in the background. The script must detect and terminate these using** **`psutil` before attempting** **`GetActiveObject`, otherwise the COM connection will freeze indefinitely.

**Python**

```
import psutil

def check_and_clear_etabs():
    """Sweeps Windows processes for ETABS and prompts a kill if needed."""
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'] and 'ETABS.exe' in proc.info['name']:
            try:
                proc.kill()
            except psutil.AccessDenied:
                raise Exception("Warning: Please close ETABS manually via Task Manager.")
```

### 3. Data Topology Mismatches (Pandas outer merge)

When comparing Model A against Model B, structural topologies often change. Pandas must handle missing elements gracefully without throwing** **`NaN` calculation errors during variance (**Δ**) calculations.

**Python**

```
import pandas as pd
import numpy as np

def safe_model_merge(df_baseline, df_optimized):
    """Merges model datasets and handles missing topology gracefully."""
    df_merged = pd.merge(
        df_baseline, df_optimized, 
        on=['Frame', 'Station', 'LoadCase'], 
        how='outer', suffixes=('_A', '_B')
    )
  
    # Calculate Variance safely
    df_merged['Variance_M3'] = np.where(
        df_merged['M3_A'].notna() & df_merged['M3_B'].notna(),
        df_merged['M3_B'] - df_merged['M3_A'],
        "Topology Changed"
    )
  
    # Fill remaining NaNs for clean Jinja rendering
    df_merged.fillna({'M3_A': "Not in Model A", 'M3_B': "Not in Model B"}, inplace=True)
    return df_merged
```

### 4. Safe Thread Routing

To keep the CustomTkinter GUI responsive, extraction runs on a background thread. Exceptions must be routed back to the main GUI via a thread-safe** **`queue.Queue()`, ensuring the user receives human-readable alerts rather than a frozen progress bar.


## 🔀 Part 2: Mock Data Generator (Stub Environment)

The ETABS COM API is locked to the Windows OS. To allow continuous UI/UX and Word template development on macOS or Linux, the software automatically detects its environment and falls back to a locally generated, structurally plausible dataset.

### 1. OS & Dependency Detection Target

**Python**

```
import platform

try:
    if platform.system() != "Windows":
        raise ImportError("Non-Windows OS detected.")
    import comtypes.client
    ETABS_AVAILABLE = True
except ImportError:
    ETABS_AVAILABLE = False
    print("DEV MODE: Utilizing Mock Data Generator.")
```

### 2. Structurally Plausible Stub Generation

The mock data cannot be entirely random. It must output the exact column headers and realistic data limits (e.g., proper drift limits, plausible D/C ratios) to validate the Word template's conditional formatting.

**Python**

```
import pandas as pd
import random

class MockEtabsExtractor:
    def __init__(self, target_stories, load_combos):
        self.stories = target_stories
        self.combos = load_combos
        self.beam_sections = ["B-30x60", "B-40x80"]

    def extract_frame_forces(self, is_optimized_model=False):
        """Generates mock frame forces. Applies a variance factor for comparison testing."""
        data = []
        variance_factor = 0.85 if is_optimized_model else 1.0 

        for combo in self.combos:
            for story in self.stories:
                for i in range(1, 6):
                    data.append({
                        "Story": story,
                        "Frame": f"B{i}-{story[-1]}",
                        "SectionProperty": random.choice(self.beam_sections),
                        "LoadCase": combo,
                        "Station": "0.0",
                        "P": round(random.uniform(-10, 10) * variance_factor, 2),
                        "V2": round(random.uniform(50, 150) * variance_factor, 2),
                        "M3": round(random.uniform(100, 350) * variance_factor, 2),
                        "Status": "OK" if random.random() > 0.1 else "FAIL"
                    })
        return pd.DataFrame(data)

    def generate_comparison_test_set(self):
        """Intentionally breaks topology to test outer merge logic."""
        df_model_A = self.extract_frame_forces(is_optimized_model=False)
        df_model_B = self.extract_frame_forces(is_optimized_model=True)
      
        # Simulate an element being deleted in Model B
        df_model_B = df_model_B.iloc[2:] 
      
        return df_model_A, df_model_B
```

### 3. Execution Pipeline Integration

In the main application logic, the** **`ETABS_AVAILABLE` flag dictates the pipeline. If** **`False`, the app triggers** **`MockEtabsExtractor()`, simulates a brief** **`time.sleep()` for UI testing, and executes the** **`docxtpl` Word generation flawlessly without ever touching the COM interface.
