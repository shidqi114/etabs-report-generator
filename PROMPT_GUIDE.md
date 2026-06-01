# 🧠 CustomTkinter Frontend Development Brain & Prompt Guide

This document acts as the dedicated frontend "brain" for the **ETABS Report Generator**. It defines UI/UX engineering protocols, layout standards, thread safety guidelines, and advanced prompt templates tailored specifically for building premium, modern desktop interfaces with **CustomTkinter**.

---

## 🎨 1. Premium Visual System (Dark Mode First)

CustomTkinter is highly customizable. To ensure the interface looks bespoke, premium, and state-of-the-art:

### 🌈 Color Palette
We reject standard system gray gradients in favor of an **ultramodern slate & neon** palette:
*   **Background (Main Window):** `#0F172A` (Slate 900)
*   **Frame Background (Cards):** `#1E293B` (Slate 800)
*   **Primary Neon Blue:** `#3B82F6` (Hover: `#2563EB`)
*   **Accent/Success Green:** `#10B981` (Hover: `#059669`)
*   **Warning/Alert Orange:** `#F59E0B` (Hover: `#D97706`)
*   **Border Accents:** `#334155` (Slate 700)
*   **Text (Primary):** `#F8FAFC` (Slate 50)
*   **Text (Muted):** `#94A3B8` (Slate 400)

### 📐 Spacing & Hierarchy
*   **Consistency:** Use `padx=20` and `pady=20` as the default padding for major panels.
*   **Grouping:** Place related interactive controls into a parent `CTkFrame` with custom border widths (`border_width=1`, `border_color="#334155"`) to act as clean cards.
*   **Typography:** Use sans-serif styling (`font=("Inter", 12)` or `font=("Helvetica Neue", 12)`) and reserve bold styling exclusively for titles (`font=("Inter", 16, "bold")`).

---

## ⚙️ 2. Grid Layout & Responsiveness

To prevent components from stretching awkwardly when the application is resized, always configure grid weight rules on containers:

```python
# Configure columns to resize proportionally
self.grid_columnconfigure(0, weight=1)
self.grid_columnconfigure(1, weight=3) # Let column 1 take up more space
```

### Premium Frame Card Blueprint:
```python
class SettingsCard(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, fg_color="#1E293B", border_width=1, border_color="#334155", **kwargs)
        
        # Grid Configuration inside the Card
        self.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=2)
        
        # Title Label
        self.title_lbl = ctk.CTkLabel(self, text="Project Scope", font=("Inter", 14, "bold"), text_color="#F8FAFC")
        self.title_lbl.grid(row=0, column=0, columnspan=2, padx=15, pady=(15, 5), sticky="w")
```

---

## 🧵 3. Thread-Safe UI Updates (CRITICAL)

Tkinter and CustomTkinter are **not thread-safe**. Making direct changes to widgets (e.g., updating a progress bar, text label, or opening a dialog) from a background worker thread will cause random application freezes or crashes.

### The Standard Pattern:
1.  **Launch a Background Thread** for long tasks (like ETABS COM extraction or docxtpl rendering) so the GUI frame rate remains stable.
2.  **Pass Thread Updates via `.after()`**: Use `self.after(ms, callback, *args)` to safely schedule UI updates on the main thread loop.

```python
import threading
import time

def start_long_operation(self):
    # Disable UI buttons to prevent double-execution
    self.run_btn.configure(state="disabled")
    
    # Start thread
    threading.Thread(target=self.bg_worker, daemon=True).start()

def bg_worker(self):
    try:
        # 1. Start execution
        self.safe_update_status("Running database scan...", 0.3)
        time.sleep(2) # Mock computation
        
        # 2. Complete execution
        self.safe_update_status("Scan Complete!", 1.0)
    except Exception as e:
        self.safe_update_error(str(e))

# Safe dispatchers back to the main loop:
def safe_update_status(self, text, progress_val):
    self.after(0, lambda: self.status_lbl.configure(text=text))
    self.after(0, lambda: self.progress_bar.set(progress_val))

def safe_update_error(self, err_msg):
    self.after(0, lambda: messagebox.showerror("Failure", err_msg))
    self.after(0, lambda: self.run_btn.configure(state="normal"))
```

---

## 🧪 4. Cross-Platform UI Mocking

Since actual ETABS automation requires **Windows**, the application must remain perfectly editable and testable on **macOS/Linux** using responsive mockup generators.

Always encapsulate data-handling logic:
```python
try:
    import comtypes.client
    # Live ETABS COM Code
except ImportError:
    # Safe Fallback Mock Engine
    print("Non-Windows OS or library missing. Activating high-fidelity UI Mockups.")
```

---

## ⚡ 5. The "ULTRATHINK" Frontend Prompt Template

When you want the AI assistant to implement or rewrite a frontend component, copy-paste this prompt:

```markdown
ULTRATHINK
Let's redesign or build a new frontend module for this CustomTkinter application.
Specific requirement: [Describe the layout / input elements / control sliders you want]

Apply these rules strictly:
1. Core Palette: Ensure Slate 900 (#0F172A), Slate 800 (#1E293B), and Slate 700 (#334155) borders are applied.
2. Responsiveness: Use self.columnconfigure/rowconfigure grid weights so it scales cleanly when resized.
3. Typography: Strict hierarchical sizes (16px bold for Headers, 12px for body, 11px muted for captions).
4. Thread Safety: If any button triggers file I/O or background extraction, implement the safe threading/lambda .after() pattern.
5. Micro-interactions: Ensure proper hover colors on CTkButtons and live, smooth progress updates.
```
