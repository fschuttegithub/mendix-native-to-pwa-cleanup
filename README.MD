# Mendix Native-to-PWA Design Property Cleaner

## The Problem
When migrating a Mendix application from a **Native Mobile** profile to a **Progressive Web App (PWA)**, you often face a massive consistency issue.

Native layouts rely heavily on Flexbox properties stored directly on widgets (e.g., `Render children horizontal`, `Justify content`, `Spacing top`, `Align children`). When you switch these pages or snippets to a **Web** profile, Mendix no longer recognizes these properties because they don't exist in the Web theme (Atlas UI Web).

**The Result:** You are instantly hit with **thousands of "Invalid Design Property" errors**.


### Why Use This Script?
Mendix Studio Pro **does not have a built-in way** to batch-remove invalid properties.
*   You cannot select 1,200 errors and click "Delete".
*   You have to manually click every single widget and remove the property one by one.
*   For a medium-sized app, this can take hours of mindless clicking.

This tool solves that problem in **minutes**.

---

## The Solution
This is a **TypeScript automation script** using the [Mendix Model SDK](https://docs.mendix.com/apidocs-mxsdk/mxsdk/).

It scans your entire project (Pages, Snippets, and Layouts), identifies the invalid Native design properties, and either **Removes** them or **Maps** them to their Web equivalents automatically.

**Features:**
*   🚀 **Automated Cleanup**: Processes 1,000+ widgets in under a minute.
*   🧠 **Smart Mapping**: Can transform `Render children horizontal` (Native) → `Align content: Left align as a row` (Web).
*   📐 **Spacing Conversion**: Merges `Spacing top`, `Spacing left`, etc., into the correct Web `Spacing` compound property.
*   🛡️ **Safe**: Runs on a dedicated branch (recommended) so you never break your main line.

---

## Getting Started

### Prerequisites
1.  **Node.js**: Install [Node.js](https://nodejs.org/) (v16 or higher).
2.  **Mendix Personal Access Token**: Generate one in your [Mendix Developer Settings](https://user.mendix.com/).

### Installation

1.  Clone this repository.
2.  Install dependencies:
    ```bash
    npm install
    ```

### Configuration
1.  **Create the environment file**:
    Copy the provided example file `.env.example` to a new file named `.env` in the **root directory** of this project (next to `package.json`).
    ```bash
    cp .env.example .env
    ```
    *(Note: The `.env` file is ignored by Git to keep your secrets safe.)*

2.  **Edit your new `.env` file**:
    Open `.env` in a text editor and fill in your credentials:
    ```ini
    MENDIX_TOKEN=your_secret_pat_here
    MENDIX_APP_ID=your_app_id_from_mendix_sprintr
    MENDIX_BRANCH_NAME=Version_PWADesignPropertiesConversion
    ```

### Usage
Run the script:
```bash
npm run start
```
The script will:
1.  Download a working copy of your branch.
2.  Scan all modules for invalid properties.
3.  Show you a summary of what it wants to change.
4.  Ask for confirmation before committing back to the Team Server.

## Reporting
For every run, the script generates a **detailed Excel report** in the `exports/` folder (e.g., `exports/remediation-report-v2-....xlsx`).

This audit trail allows you to verify exactly what happened:
*   Which widget was modified (Document, Name, Type).
*   Which property was removed or mapped.
*   The original value vs. the new value.
*   The specific rule/reason that triggered the change.

---

## Customizing Mappings
The script's behavior is controlled entirely by `property-mappings.json`. You can modify this file to change how properties are handled.

Example rule:
```json
{
  "property": "Size",
  "value": "Maximum space",
  "elementTypes": ["DivContainer"],
  "action": "remove",
  "_reason": "Flexbox sizing property has no Web equivalent"
}
```
*   **`action`**: `"remove"` to delete checks, `"map"` to transform.
*   **`elementTypes`**: Apply rules only to specific widgets (e.g., `["DivContainer", "ListView"]`).


