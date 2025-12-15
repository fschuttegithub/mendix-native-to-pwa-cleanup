/**
 * Mendix Native-to-Web Migration Remediation Script v2.0
 * 
 * PURPOSE: Detect and remediate invalid Native design properties from Mendix models.
 * 
 * KEY IMPROVEMENTS (v2.0):
 * - Element-type aware mappings: Only applies mappings to appropriate widget types
 * - Action-based processing: Supports "map" (transform) and "remove" (delete) actions
 * - Validates against web design properties before applying
 * - Better spacing handling with direction and type (margin/padding)
 * - Detailed logging and Excel export with action taken
 * 
 * Author: Mendix Expert Implementation
 */

import { MendixPlatformClient, OnlineWorkingCopy, setPlatformConfig } from "mendixplatformsdk";
import { domainmodels, pages, type IModel, projects, texts } from "mendixmodelsdk";
import * as fs from "fs";
import * as path from "path";
import * as XLSX from "xlsx";
import * as readline from "readline";

// =============================================================================
// CONFIGURATION SECTION
// =============================================================================

// The unique App ID from the Mendix Portal (General Settings)
// The unique App ID from the Mendix Portal (General Settings)
const APP_ID = process.env.MENDIX_APP_ID || "";

// The branch to modify. Ideally, perform this on a dedicated 'migration' branch.
const BRANCH_NAME = process.env.MENDIX_BRANCH_NAME || "Version_PWADesignPropertiesConversion";

// Personal Access Token (PAT).
const MENDIX_TOKEN = process.env.MENDIX_TOKEN || "";

// =============================================================================
// TYPE DEFINITIONS
// =============================================================================

interface PropertyMappingV2 {
    property: string;
    value: string;
    elementTypes: string[];  // ["*"] means all elements, or specific types
    action: "map" | "remove" | "mapToWidgetProperty";
    mappedProperty?: string;
    mappedValue?: string;
    mappedDirection?: string;  // For spacing: top/bottom/left/right
    mappedType?: string;       // For spacing: margin/padding
    widgetProperty?: string;   // For mapToWidgetProperty: the widget property name
    widgetValue?: string;      // For mapToWidgetProperty: the value to set
    _reason?: string;
    _comment?: string;
}

interface SpacingState {
    margin: { top: string; right: string; bottom: string; left: string };
    padding: { top: string; right: string; bottom: string; left: string };
}

interface MappingsFileV2 {
    _metadata: {
        version: string;
        description: string;
        created: string;
        notes: string[];
    };
    mappings: PropertyMappingV2[];
}

interface ProcessingResult {
    widgets: number;
    props: number;
    mapped: number;
    removed: number;
    widgetPropsMapped: number;  // For mapToWidgetProperty actions
    skipped: number;
    skippedNativeLayouts?: number;
}

interface CsvDataRow {
    element: string;
    elementType: string;
    document: string;
    module: string;
    property: string;
    value: string;
    action: "mapped" | "removed" | "skipped" | "widgetPropertySet";
    mappedProperty?: string;
    mappedValue?: string;
    reason?: string;
}

// =============================================================================
// GLOBAL DATA STRUCTURES
// =============================================================================

let MAPPINGS_DATA: PropertyMappingV2[] = [];

// Lookup map: "property|value|elementType" -> mapping
// Also supports wildcard: "property|value|*"
const MAPPINGS_LOOKUP = new Map<string, PropertyMappingV2>();

// Set of all property names that have mappings (for quick filtering)
const MAPPED_PROPERTY_NAMES = new Set<string>();

// Target element types to process (only these will be modified)
const TARGET_ELEMENT_TYPES = new Set([
    "DynamicText",
    "DivContainer",
    "StaticImageViewer",
    "ListView",
    "ActionButton",
    "ReferenceSelector",
    "DropDown",
    "InputReferenceSetSelector",
    "CheckBox",
    "com.mendix.widget.native.badge.Badge",
    "TextArea",
    "ImageViewer",
    "Text"
]);

// =============================================================================
// INITIALIZATION
// =============================================================================

function loadMappings(): boolean {
    try {
        const mappingsPath = path.join(process.cwd(), "property-mappings.json");

        if (!fs.existsSync(mappingsPath)) {
            console.error("ERROR: property-mappings.json not found!");
            console.error("Please ensure the file exists in the project root.");
            return false;
        }

        const mappingsContent = fs.readFileSync(mappingsPath, "utf-8");
        const mappingsFile = JSON.parse(mappingsContent) as MappingsFileV2;

        // Filter out comment-only entries
        MAPPINGS_DATA = mappingsFile.mappings.filter(m => m.property && m.value);

        // Build lookup structures
        for (const mapping of MAPPINGS_DATA) {
            MAPPED_PROPERTY_NAMES.add(mapping.property);

            // Create lookup keys for each element type
            for (const elementType of mapping.elementTypes) {
                const key = `${mapping.property}|${mapping.value}|${elementType}`;
                MAPPINGS_LOOKUP.set(key, mapping);
            }
        }

        const mapCount = MAPPINGS_DATA.filter(m => m.action === "map").length;
        const removeCount = MAPPINGS_DATA.filter(m => m.action === "remove").length;

        console.log(`> Loaded ${MAPPINGS_DATA.length} property mappings from property-mappings.json`);
        console.log(`  - Map actions: ${mapCount}`);
        console.log(`  - Remove actions: ${removeCount}`);
        console.log(`  - Unique properties: ${MAPPED_PROPERTY_NAMES.size}`);

        return true;
    } catch (error) {
        console.error("ERROR: Could not load property-mappings.json:", error);
        return false;
    }
}

/**
 * Find a mapping for a given property/value/elementType combination
 * Checks specific element type first, then falls back to wildcard (*)
 */
function findMapping(property: string, value: string, elementType: string): PropertyMappingV2 | undefined {
    // First try specific element type
    const specificKey = `${property}|${value}|${elementType}`;
    let mapping = MAPPINGS_LOOKUP.get(specificKey);

    if (mapping) return mapping;

    // Try wildcard
    const wildcardKey = `${property}|${value}|*`;
    mapping = MAPPINGS_LOOKUP.get(wildcardKey);

    return mapping;
}

// =============================================================================
// USER CONFIRMATION HELPER
// =============================================================================

/**
 * Prompt user for confirmation
 */
function askConfirmation(question: string): Promise<boolean> {
    const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout
    });

    return new Promise((resolve) => {
        rl.question(`${question} (yes/no): `, (answer) => {
            rl.close();
            const normalized = answer.trim().toLowerCase();
            resolve(normalized === 'yes' || normalized === 'y');
        });
    });
}

// =============================================================================
// MAIN EXECUTION FLOW
// =============================================================================

async function main() {
    console.log(`\n${"=".repeat(70)}`);
    console.log(`  MENDIX NATIVE-TO-WEB DESIGN PROPERTY REMEDIATION v2.0`);
    console.log(`${"=".repeat(70)}\n`);

    // Load mappings
    if (!loadMappings()) {
        process.exit(1);
    }

    if (!MENDIX_TOKEN || !APP_ID) {
        console.error("CRITICAL ERROR: Configuration missing.");
        console.error("Please set MENDIX_TOKEN and MENDIX_APP_ID environment variables.");
        return;
    }

    // Configure Platform SDK
    setPlatformConfig({
        mendixToken: MENDIX_TOKEN
    });

    const client = new MendixPlatformClient();

    console.log(`\n> Target App ID: ${APP_ID}`);
    console.log(`> Target Branch: ${BRANCH_NAME}`);
    console.log(`> Mode: ACTIVE - Properties will be modified`);
    console.log(`> Target Elements: ${Array.from(TARGET_ELEMENT_TYPES).join(', ')}\n`);

    try {
        const app = client.getApp(APP_ID);

        console.log(`> Opening working copy...`);
        const workingCopy = await app.createTemporaryWorkingCopy(BRANCH_NAME);
        const model = await workingCopy.openModel();

        console.log(`> Model loaded successfully.`);
        console.log(`> Starting analysis of UI documents...\n`);

        // Process all modules
        const allModules = model.allModules();
        const totals: ProcessingResult = {
            widgets: 0,
            props: 0,
            mapped: 0,
            removed: 0,
            widgetPropsMapped: 0,
            skipped: 0,
            skippedNativeLayouts: 0
        };

        const csvData: CsvDataRow[] = [];

        for (const module of allModules) {
            if (module.name === "System") continue;

            console.log(`\n${"─".repeat(50)}`);
            console.log(`Processing Module: [${module.name}]`);
            console.log(`${"─".repeat(50)}`);

            // Process Pages
            const pageDocuments = model.allPages().filter(
                (p) => p.qualifiedName?.startsWith(module.name) ?? false
            );
            const pagesResult = await processDocumentCollection(
                pageDocuments, "Page", module.name, csvData
            );

            // Process Snippets
            const snippetDocuments = model.allSnippets().filter(
                (s) => s.qualifiedName?.startsWith(module.name) ?? false
            );
            const snippetsResult = await processDocumentCollection(
                snippetDocuments, "Snippet", module.name, csvData
            );

            // Process Layouts
            const layoutDocuments = model.allLayouts().filter(
                (l) => l.qualifiedName?.startsWith(module.name) ?? false
            );
            const layoutsResult = await processDocumentCollection(
                layoutDocuments, "Layout", module.name, csvData
            );

            // Aggregate stats
            totals.widgets += pagesResult.widgets + snippetsResult.widgets + layoutsResult.widgets;
            totals.props += pagesResult.props + snippetsResult.props + layoutsResult.props;
            totals.mapped += pagesResult.mapped + snippetsResult.mapped + layoutsResult.mapped;
            totals.removed += pagesResult.removed + snippetsResult.removed + layoutsResult.removed;
            totals.widgetPropsMapped += pagesResult.widgetPropsMapped + snippetsResult.widgetPropsMapped + layoutsResult.widgetPropsMapped;
            totals.skipped += pagesResult.skipped + snippetsResult.skipped + layoutsResult.skipped;
            totals.skippedNativeLayouts! += pagesResult.skippedNativeLayouts || 0;
        }

        // Write Excel report
        if (csvData.length > 0) {
            writeExcelReport(csvData);
        }

        // Summary
        console.log(`\n${"=".repeat(70)}`);
        console.log(`  REMEDIATION COMPLETE`);
        console.log(`${"=".repeat(70)}`);
        console.log(`  Total Widgets Processed: ${totals.widgets}`);
        console.log(`  Total Properties Found:  ${totals.props}`);
        console.log(`  ├─ Mapped (design props):  ${totals.mapped}`);
        console.log(`  ├─ Mapped (widget props):  ${totals.widgetPropsMapped}`);
        console.log(`  ├─ Removed:                ${totals.removed}`);
        console.log(`  └─ Skipped (no mapping):   ${totals.skipped}`);

        if (totals.skippedNativeLayouts! > 0) {
            console.log(`  Native Layout Pages:     ${totals.skippedNativeLayouts} (skipped)`);
        }

        const totalModified = totals.mapped + totals.removed + totals.widgetPropsMapped;

        if (totalModified > 0) {
            console.log(`\n> Modified ${totalModified} properties across ${totals.widgets} widgets.`);
            console.log(`\n${"=".repeat(70)}`);
            console.log(`  COMMIT SUMMARY`);
            console.log(`${"=".repeat(70)}`);
            console.log(`  Branch:        ${BRANCH_NAME}`);
            console.log(`  Widgets:        ${totals.widgets}`);
            console.log(`  Properties:`);
            console.log(`    - Mapped:     ${totals.mapped} design properties`);
            console.log(`    - Widget props: ${totals.widgetPropsMapped} widget properties`);
            console.log(`    - Removed:    ${totals.removed} properties`);
            console.log(`    - Skipped:    ${totals.skipped} (no mapping)`);
            console.log(`  Total changes: ${totalModified} properties`);
            console.log(`${"=".repeat(70)}\n`);

            // Ask for user confirmation
            const shouldCommit = await askConfirmation(
                `\n⚠️  Do you want to commit these changes to branch "${BRANCH_NAME}"?`
            );

            if (!shouldCommit) {
                console.log(`\n> Commit cancelled by user. Changes have NOT been saved.`);
                console.log(`> Working copy will be discarded.`);
                return;
            }

            console.log(`\n> Flushing changes...`);
            model.flushChanges();

            console.log(`> Committing to Team Server...`);
            await workingCopy.commitToRepository(BRANCH_NAME, {
                commitMessage: `Remediated ${totalModified} Native design properties (${totals.mapped} design props mapped, ${totals.widgetPropsMapped} widget props set, ${totals.removed} removed)`
            });
            console.log(`> Success! Changes committed.`);
        } else {
            console.log(`\n> No properties were modified.`);
        }

    } catch (error) {
        console.error("CRITICAL FAILURE:", error);
    }
}

// =============================================================================
// DOCUMENT PROCESSING
// =============================================================================

async function processDocumentCollection(
    documents: any,
    typeName: string,
    moduleName: string,
    csvData: CsvDataRow[]
): Promise<ProcessingResult> {
    const result: ProcessingResult = {
        widgets: 0,
        props: 0,
        mapped: 0,
        removed: 0,
        widgetPropsMapped: 0,
        skipped: 0,
        skippedNativeLayouts: 0
    };

    for (const docInfo of documents) {
        try {
            const document = await docInfo.load();
            const docName = docInfo.qualifiedName || docInfo.name || 'Unknown';

            // Skip Native layout pages
            // Note: For Layout documents, 'document' is the Layout itself.
            if (typeName === "Page" && document instanceof pages.Page) {
                const page = document;
                const layoutCall = page.layoutCall;

                if (layoutCall?.layout) {
                    const layout = layoutCall.layout;
                    const layoutName = layout.qualifiedName?.toLowerCase() || "";

                    if (layoutName.includes("native")) {
                        result.skippedNativeLayouts!++;
                        console.log(`  [SKIP] ${docName} - Native layout`);
                        continue;
                    }
                }

                // Process Page design properties - remove all of them (Pages don't support design properties in Web)
                if (page.appearance?.designProperties?.length) {
                    const pageProps = page.appearance.designProperties;
                    const propsToRemove = [...pageProps]; // Copy array to avoid modification during iteration

                    if (propsToRemove.length > 0) {
                        console.log(`\n  [Page] ${docName}`);
                        console.log(`    └─ Page: ${docName}`);
                        console.log(`       ✗ Removed: ${propsToRemove.length} design properties from Page`);

                        for (const prop of propsToRemove) {
                            const propKey = prop.key || 'Unknown';
                            const propValue = extractPropertyValue(prop);

                            console.log(`         - ${propKey}`);

                            // Record in CSV
                            csvData.push({
                                element: docName,
                                elementType: "Page",
                                document: docName,
                                module: moduleName,
                                property: propKey,
                                value: propValue,
                                action: "removed",
                                reason: "Pages do not support design properties in Web profile"
                            });

                            // Remove the property
                            try {
                                if (pageProps.indexOf(prop) >= 0) {
                                    pageProps.remove(prop);
                                }
                            } catch (error: any) {
                                console.warn(`       ⚠ Failed to remove Page property "${propKey}": ${error?.message || String(error)}`);
                            }
                        }

                        result.props += propsToRemove.length;
                        result.removed += propsToRemove.length;
                    }
                }
            }

            // Traverse document
            try {
                document.traverse((structure: any) => {
                    try {
                        if (structure instanceof pages.Widget) {
                            const widgetResult = processWidget(
                                structure, docName, typeName, moduleName, csvData
                            );

                            if (widgetResult.props > 0) {
                                result.widgets++;
                                result.props += widgetResult.props;
                                result.mapped += widgetResult.mapped;
                                result.removed += widgetResult.removed;
                                result.skipped += widgetResult.skipped;
                            }
                        }
                    } catch (widgetError) {
                        console.warn(`  Warning: Error processing widget in ${docName}:`, widgetError);
                    }
                });
            } catch (traverseError) {
                console.warn(`  Warning: Error traversing ${docName}:`, traverseError);
            }

        } catch (e) {
            console.warn(`  Warning: Could not process '${docInfo.qualifiedName}':`, e);
        }
    }

    return result;
}

// =============================================================================
// WIDGET PROCESSING
// =============================================================================

function processWidget(
    widget: pages.Widget,
    docName: string,
    docType: string,
    moduleName: string,
    csvData: CsvDataRow[]
): ProcessingResult {
    const result: ProcessingResult = {
        widgets: 0,
        props: 0,
        mapped: 0,
        removed: 0,
        widgetPropsMapped: 0,
        skipped: 0
    };

    if (!widget.appearance?.designProperties?.length) {
        return result;
    }

    const widgetName = (widget as any).name || 'Unnamed';
    const widgetType = getWidgetTypeName(widget);

    // Skip widgets that are not in our target element types
    if (!TARGET_ELEMENT_TYPES.has(widgetType)) {
        return result;
    }

    const propertiesList = widget.appearance.designProperties;

    // Track properties to modify, remove, and widget properties to set
    const toModify: Array<{ prop: pages.DesignPropertyValue, mapping: PropertyMappingV2 }> = [];
    const toRemove: pages.DesignPropertyValue[] = [];
    const toSetWidgetProp: Array<{ prop: pages.DesignPropertyValue, mapping: PropertyMappingV2 }> = [];
    const skippedProps: Array<{ prop: string, value: string, reason: string }> = [];

    // Track target properties that have already been mapped to prevent duplicates (e.g., multiple native props mapping to "Align content")
    const mappedTargetProperties = new Set<string>();

    // Track spacing properties for later aggregation into a single "Spacing" design property
    const spacingPropertiesToMap: Array<{
        prop: pages.DesignPropertyValue;
        mapping: PropertyMappingV2;
        direction: string;
        spacingType: "margin" | "padding";
    }> = [];

    // Check if widget already has a "Spacing" property (will be reused as the target)
    const existingSpacingProp = propertiesList.find((p) => p.key === "Spacing");

    // Analyze all properties - first pass: collect spacing properties (do not map yet)
    for (const prop of propertiesList) {
        const propKey = prop.key;
        if (!propKey) continue;

        // Skip if this property name isn't in our mappings
        if (!MAPPED_PROPERTY_NAMES.has(propKey)) continue;

        // Extract value
        let propValue = extractPropertyValue(prop);
        if (!propValue) continue;

        result.props++;

        // Find mapping for this property/value/elementType
        const mapping = findMapping(propKey, propValue, widgetType);

        if (!mapping) {
            // No mapping found - skip
            skippedProps.push({
                prop: propKey,
                value: propValue,
                reason: "No mapping defined for this property/value/element combination"
            });
            result.skipped++;
            continue;
        }

        if (mapping.action === "remove") {
            toRemove.push(prop);
            result.removed++;

            csvData.push({
                element: widgetName,
                elementType: widgetType,
                document: docName,
                module: moduleName,
                property: propKey,
                value: propValue,
                action: "removed",
                reason: mapping._reason || "No Web equivalent"
            });
        } else if (mapping.action === "map") {
            // Special handling for Spacing properties: collect for aggregation
            if (mapping.mappedProperty === "Spacing" && mapping.mappedDirection) {
                const spacingType: "margin" | "padding" = mapping.mappedType === "padding" ? "padding" : "margin";
                spacingPropertiesToMap.push({
                    prop,
                    mapping,
                    direction: mapping.mappedDirection,
                    spacingType
                });
            } else {
                // Non-spacing property - normal mapping
                const targetProp = mapping.mappedProperty || "";

                // CHECK FOR DUPLICATES: If this target property has already been mapped by a previous native property on this widget, skip it.
                // This handles cases like "Render children horizontal" AND "Justify content" both mapping to "Align content".
                // The first one encountered "wins".
                if (targetProp && mappedTargetProperties.has(targetProp)) {
                    // ACTION: Remove the conflicting property to avoid "Invalid Design Property" errors.
                    toRemove.push(prop);
                    result.removed++;

                    csvData.push({
                        element: widgetName,
                        elementType: widgetType,
                        document: docName,
                        module: moduleName,
                        property: propKey,
                        value: propValue,
                        action: "removed",
                        reason: `Duplicate target '${targetProp}' - collision with prioritized mapping`
                    });
                    continue;
                }

                if (targetProp) {
                    mappedTargetProperties.add(targetProp);
                }

                toModify.push({ prop, mapping });
                result.mapped++;

                csvData.push({
                    element: widgetName,
                    elementType: widgetType,
                    document: docName,
                    module: moduleName,
                    property: propKey,
                    value: propValue,
                    action: "mapped",
                    mappedProperty: mapping.mappedProperty || "",
                    mappedValue: mapping.mappedValue || ""
                });
            }
        } else if (mapping.action === "mapToWidgetProperty") {
            toSetWidgetProp.push({ prop, mapping });
            result.widgetPropsMapped++;

            csvData.push({
                element: widgetName,
                elementType: widgetType,
                document: docName,
                module: moduleName,
                property: propKey,
                value: propValue,
                action: "widgetPropertySet",
                mappedProperty: mapping.widgetProperty || "",
                mappedValue: mapping.widgetValue || "",
                reason: mapping._reason || "Mapped to widget property"
            });
        }
    }

    // Process collected spacing properties:
    // Merge all directions into a single "Spacing" design property (margin/padding per side).
    if (spacingPropertiesToMap.length > 0) {
        // 1. Calculate the final merged state (Native + Existing Web)
        const baseState: SpacingState = {
            margin: { top: "None", right: "None", bottom: "None", left: "None" },
            padding: { top: "None", right: "None", bottom: "None", left: "None" }
        };

        // If widget already has Spacing, parse it first to preserve existing web adjustments
        if (existingSpacingProp) {
            const parsed = parseSpacingState(existingSpacingProp);
            if (parsed) {
                Object.assign(baseState.margin, parsed.margin);
                Object.assign(baseState.padding, parsed.padding);
            }
        }

        // Overwrite with mapped Native values
        for (const item of spacingPropertiesToMap) {
            const dir = item.direction || "top";
            const type = item.spacingType;
            const val = item.mapping.mappedValue || "None";
            const oldVal = baseState[type][dir as keyof typeof baseState.margin];
            baseState[type][dir as keyof typeof baseState.margin] = val;
            console.log(`       → Mapping: "${item.prop.key}"="${extractPropertyValue(item.prop)}" → Spacing ${type}.${dir}="${val}"${oldVal !== val ? ` (was: ${oldVal})` : ''}`);
        }

        // 2. Create the NEW completely valid SDK structure
        // Use createInAppearanceUnderDesignProperties to properly attach to widget
        const newSpacingProp = pages.DesignPropertyValue.createInAppearanceUnderDesignProperties(widget.appearance);
        newSpacingProp.key = "Spacing";

        // Create the compound value and attach it
        const compoundValue = pages.CompoundDesignPropertyValue.createIn(newSpacingProp);
        newSpacingProp.value = compoundValue;

        // Helper to add inner properties using createIn methods
        const addInnerProp = (key: string, val: string) => {
            if (val === "None") return;
            const innerProp = pages.DesignPropertyValue.createInCompoundDesignPropertyValueUnderProperties(compoundValue);
            innerProp.key = key;
            const optionVal = pages.OptionDesignPropertyValue.createIn(innerProp);
            optionVal.option = val;
            innerProp.value = optionVal;
        };

        // Populate based on merged state
        addInnerProp("margin-top", baseState.margin.top);
        addInnerProp("margin-right", baseState.margin.right);
        addInnerProp("margin-bottom", baseState.margin.bottom);
        addInnerProp("margin-left", baseState.margin.left);
        addInnerProp("padding-top", baseState.padding.top);
        addInnerProp("padding-right", baseState.padding.right);
        addInnerProp("padding-bottom", baseState.padding.bottom);
        addInnerProp("padding-left", baseState.padding.left);

        // 3. Remove existing Spacing property if present (to avoid duplicates)
        if (existingSpacingProp) {
            propertiesList.remove(existingSpacingProp);
        }

        // Note: Property is already added by createInAppearanceUnderDesignProperties

        // 5. Mark Native spacing properties for removal (they're being merged into Spacing)
        // Note: These are tracked separately so they show up as "mapped" not "removed" in logs
        const spacingPropsToRemove: pages.DesignPropertyValue[] = [];
        for (const item of spacingPropertiesToMap) {
            spacingPropsToRemove.push(item.prop);
            // Don't add to toRemove - we'll handle removal separately and log them as mapped
        }

        // Count each individual spacing property that was mapped
        // Each spacing property (top, bottom, left, right) counts as a mapped property
        result.mapped += spacingPropertiesToMap.length;

        // Record CSV entries for each mapped spacing item
        for (const item of spacingPropertiesToMap) {
            csvData.push({
                element: widgetName,
                elementType: widgetType,
                document: docName,
                module: moduleName,
                property: item.prop.key,
                value: extractPropertyValue(item.prop),
                action: "mapped",
                mappedProperty: "Spacing",
                mappedValue: `${item.spacingType}.${item.direction}=${item.mapping.mappedValue || ""}`
            });
        }
    }

    // Log if we found anything
    if (result.props > 0) {
        console.log(`\n  [${docType}] ${docName}`);
        console.log(`    └─ ${widgetType}: ${widgetName}`);

        if (result.mapped > 0) {
            console.log(`       ✓ Mapped: ${result.mapped} properties`);
            // Show non-spacing mappings
            for (const item of toModify) {
                const m = item.mapping;
                if (m.mappedDirection && m.mappedType) {
                    console.log(`         - ${item.prop.key} → ${m.mappedProperty} ${m.mappedType}.${m.mappedDirection}="${m.mappedValue}"`);
                } else {
                    console.log(`         - ${item.prop.key} → ${m.mappedProperty}="${m.mappedValue}"`);
                }
            }
            // Show spacing mappings (these are being merged into Spacing property)
            if (spacingPropertiesToMap.length > 0) {
                for (const item of spacingPropertiesToMap) {
                    console.log(`         - ${item.prop.key} → Spacing ${item.spacingType}.${item.direction}="${item.mapping.mappedValue || ""}"`);
                }
            }
        }

        if (result.removed > 0) {
            console.log(`       ✗ Removed: ${result.removed} properties`);
            // Only show properties that are actually removed (not spacing properties which are mapped)
            const actuallyRemoved = toRemove.filter(prop =>
                !spacingPropertiesToMap.some(sp => sp.prop === prop)
            );
            for (const prop of actuallyRemoved) {
                console.log(`         - ${prop.key}`);
            }
        }

        if (result.widgetPropsMapped > 0) {
            console.log(`       ⚙ Widget props set: ${result.widgetPropsMapped}`);
            for (const item of toSetWidgetProp) {
                console.log(`         - ${item.prop.key} → widget.${item.mapping.widgetProperty}="${item.mapping.widgetValue}"`);
            }
        }

        if (result.skipped > 0) {
            console.log(`       ○ Skipped: ${result.skipped} properties (no mapping)`);
        }
    }

    // Execute modifications
    for (const item of toModify) {
        try {
            applyMapping(item.prop, item.mapping);
        } catch (error: any) {
            const propKey = item.prop.key || 'Unknown';
            const propValue = extractPropertyValue(item.prop);
            console.warn(`       ⚠ Failed to modify property "${propKey}"="${propValue}": ${error?.message || String(error)}`);
        }
    }

    // Execute removals (including spacing properties that were merged)
    // Combine both regular removals and spacing properties that need to be removed
    const allPropsToRemove = [...toRemove];
    for (const item of spacingPropertiesToMap) {
        if (!allPropsToRemove.includes(item.prop)) {
            allPropsToRemove.push(item.prop);
        }
    }

    for (const prop of allPropsToRemove) {
        try {
            // Use IList.remove() method instead of splice() for proper SDK handling
            if (propertiesList.indexOf(prop) >= 0) {
                propertiesList.remove(prop);
            }
        } catch (error: any) {
            console.warn(`       ⚠ Failed to remove property: ${error?.message || String(error)}`);
        }
    }

    // Execute widget property changes (for Button style -> buttonStyle)
    for (const item of toSetWidgetProp) {
        try {
            applyWidgetPropertyMapping(widget, item.mapping);
            // Also remove the design property since we've mapped it to widget property
            // Use IList.remove() method instead of splice() for proper SDK handling
            if (propertiesList.indexOf(item.prop) >= 0) {
                propertiesList.remove(item.prop);
            }
        } catch (error: any) {
            console.warn(`       ⚠ Failed to set widget property: ${error?.message || String(error)}`);
        }
    }

    return result;
}

// =============================================================================
// HELPER FUNCTIONS
// =============================================================================

/**
 * Get the widget type name for lookup
 */
function getWidgetTypeName(widget: pages.Widget): string {
    // Try to get the widget's structural type name
    const structTypeName = (widget as any).structureTypeName;

    if (structTypeName) {
        // Convert from SDK format (e.g., "Pages$ActionButton") to simple name
        const parts = structTypeName.split("$");
        if (parts.length > 1) {
            return parts[1];
        }
        return structTypeName;
    }

    // Fallback: try widgetId for pluggable widgets
    const widgetId = (widget as any).widgetId;
    if (widgetId) {
        return widgetId;
    }

    // Last resort: constructor name
    return widget.constructor.name;
}

/**
 * Extract the value from a design property
 * Uses instanceof checks for type safety instead of as any
 */
function extractPropertyValue(prop: pages.DesignPropertyValue): string {
    const value = prop.value;

    if (value === undefined || value === null) {
        return '';
    }

    // Check if it's an OptionDesignPropertyValue
    if (value instanceof pages.OptionDesignPropertyValue) {
        return value.option || '';
    }

    // Check if it's a CompoundDesignPropertyValue (for Spacing, etc.)
    if (value instanceof pages.CompoundDesignPropertyValue) {
        // For compound values, return a descriptive string
        return '[Compound]';
    }

    // Check if it's a ToggleDesignPropertyValue
    // The presence of the object itself indicates the toggle is "on" (true)
    if (value instanceof pages.ToggleDesignPropertyValue) {
        return "true";
    }

    // Handle primitive types
    if (typeof value === 'boolean') {
        return String(value);
    }

    if (typeof value === 'number') {
        return String(value);
    }

    if (typeof value === 'string') {
        return value;
    }

    // Fallback for unknown types
    return String(value);
}

/**
 * Apply a mapping to a property (modify in place)
 * Uses SDK factory methods to create proper model elements instead of plain objects.
 */
function applyMapping(prop: pages.DesignPropertyValue, mapping: PropertyMappingV2): void {
    if (!mapping.mappedProperty || !mapping.mappedValue) return;

    const model = prop.model; // Access model reference from the element itself

    // 1. Update Key
    prop.key = mapping.mappedProperty;

    // 2. Update Value - Create a NEW Option value
    // We replace the value entirely to ensure it's the correct type (OptionDesignPropertyValue)
    // This handles cases where the old value might have been a simple string or different type.

    const newValue = pages.OptionDesignPropertyValue.create(model);
    newValue.option = mapping.mappedValue;

    // Assign the new model element to the property
    prop.value = newValue;
}

/**
 * Apply a widget property mapping (e.g., Button style -> buttonStyle widget property)
 * This sets the actual widget property, not a design property
 */
function applyWidgetPropertyMapping(widget: pages.Widget, mapping: PropertyMappingV2): void {
    if (!mapping.widgetProperty || !mapping.widgetValue) {
        throw new Error("Widget property mapping missing widgetProperty or widgetValue");
    }

    // Handle buttonStyle specifically for ActionButton
    if (mapping.widgetProperty === "buttonStyle" && widget instanceof pages.ActionButton) {
        // Map the string value to the ButtonStyle enum
        const buttonStyleMap: { [key: string]: pages.ButtonStyle } = {
            "Default": pages.ButtonStyle.Default,
            "Inverse": pages.ButtonStyle.Inverse,
            "Primary": pages.ButtonStyle.Primary,
            "Info": pages.ButtonStyle.Info,
            "Success": pages.ButtonStyle.Success,
            "Warning": pages.ButtonStyle.Warning,
            "Danger": pages.ButtonStyle.Danger
        };

        const newStyle = buttonStyleMap[mapping.widgetValue];
        if (newStyle) {
            widget.buttonStyle = newStyle;
        } else {
            throw new Error(`Unknown button style value: ${mapping.widgetValue}`);
        }
    } else {
        // Generic fallback for other widget properties (future use)
        (widget as any)[mapping.widgetProperty] = mapping.widgetValue;
    }
}

/**
 * Create a complete Spacing design property from scratch using SDK model elements.
 * This builds the entire nested structure correctly using the SDK factory methods.
 * NOTE: This function is deprecated - use inline creation with createIn methods instead.
 * Kept for reference but no longer used.
 */
function createSpacingProperty(model: IModel, spacingState: SpacingState): pages.DesignPropertyValue {
    // DEPRECATED: This function is no longer used.
    // Use createInAppearanceUnderDesignProperties and createInCompoundDesignPropertyValueUnderProperties instead.
    throw new Error("createSpacingProperty is deprecated. Use createIn methods directly.");
}

/**
 * Parse an existing Spacing design property into our SpacingState structure.
 * Handles CompoundDesignPropertyValue with properties array.
 */
function parseSpacingState(prop: pages.DesignPropertyValue | undefined): SpacingState | undefined {
    if (!prop) return undefined;
    const val = prop.value;
    if (!val) return undefined;

    // Initialize default state
    const margin = { top: "None", right: "None", bottom: "None", left: "None" };
    const padding = { top: "None", right: "None", bottom: "None", left: "None" };

    // Check if this is a CompoundDesignPropertyValue with a properties array
    if (val instanceof pages.CompoundDesignPropertyValue) {
        // Parse properties array format: [{ key: "margin-top", value: OptionDesignPropertyValue }, ...]
        for (const propEntry of val.properties) {
            const key = propEntry.key;
            const propValue = propEntry.value;

            if (!key || !propValue) continue;

            // Extract option value if it's an OptionDesignPropertyValue
            let option: string | undefined;
            if (propValue instanceof pages.OptionDesignPropertyValue) {
                option = propValue.option;
            }

            if (!option) continue;

            // Parse keys like "margin-top", "padding-left", etc.
            const parts = key.split("-");
            if (parts.length === 2) {
                const type = parts[0]; // "margin" or "padding"
                const direction = parts[1]; // "top", "right", "bottom", "left"

                if ((type === "margin" || type === "padding") &&
                    (direction === "top" || direction === "right" || direction === "bottom" || direction === "left")) {
                    (type === "margin" ? margin : padding)[direction as keyof typeof margin] = option;
                }
            }
        }
        return { margin, padding };
    }

    return undefined;
}

/**
 * Write Excel report
 */
function writeExcelReport(csvData: CsvDataRow[]): void {
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, -5);
    const filename = `remediation-report-v2-${timestamp}.xlsx`;

    const exportsDir = path.join(process.cwd(), 'exports');
    if (!fs.existsSync(exportsDir)) {
        fs.mkdirSync(exportsDir, { recursive: true });
    }

    const filepath = path.join(exportsDir, filename);

    // Prepare worksheet data
    const headers = [
        'Element', 'Element Type', 'Document', 'Module',
        'Property', 'Value', 'Action', 'Mapped Property', 'Mapped Value', 'Reason'
    ];

    const worksheetData = [
        headers,
        ...csvData.map(row => [
            row.element,
            row.elementType,
            row.document,
            row.module,
            row.property,
            row.value,
            row.action,
            row.mappedProperty || '',
            row.mappedValue || '',
            row.reason || ''
        ])
    ];

    // Create workbook
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);

    // Column widths
    worksheet['!cols'] = [
        { wch: 25 }, // element
        { wch: 35 }, // elementType
        { wch: 40 }, // document
        { wch: 20 }, // module
        { wch: 25 }, // property
        { wch: 20 }, // value
        { wch: 10 }, // action
        { wch: 25 }, // mappedProperty
        { wch: 20 }, // mappedValue
        { wch: 50 }  // reason
    ];

    XLSX.utils.book_append_sheet(workbook, worksheet, 'Remediation Report');

    // Summary sheet
    const summaryData = [
        ['Summary', ''],
        ['Generated', new Date().toISOString()],
        ['Total Properties', csvData.length.toString()],
        ['Mapped', csvData.filter(r => r.action === 'mapped').length.toString()],
        ['Removed', csvData.filter(r => r.action === 'removed').length.toString()],
        ['Skipped', csvData.filter(r => r.action === 'skipped').length.toString()]
    ];

    const summarySheet = XLSX.utils.aoa_to_sheet(summaryData);
    XLSX.utils.book_append_sheet(workbook, summarySheet, 'Summary');

    XLSX.writeFile(workbook, filepath);

    console.log(`\n> Excel report: exports/${filename}`);
    console.log(`  Total entries: ${csvData.length}`);
}

// =============================================================================
// EXECUTION
// =============================================================================

main().catch((error) => {
    console.error("Unhandled error:", error);
    process.exit(1);
});

process.on('unhandledRejection', (reason, promise) => {
    console.error('Unhandled Rejection:', reason);
    process.exit(1);
});

