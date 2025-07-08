import React, { useState, useEffect } from 'react';
import { Stack, Text, MessageBar, MessageBarType, Spinner, SpinnerSize, PrimaryButton } from '@fluentui/react';
import { CompoundSelector } from './CompoundSelector';
import { TemplateSelector } from './TemplateSelector';
import { AppState, Compound, Template, BatchData, API_BASE_URL } from '../../types';
import axios from 'axios';
import '../taskpane.css';

export const App: React.FC = () => {
    const [state, setState] = useState<AppState>({
        extractedData: [],
        isLoading: false
    });

    const [compounds, setCompounds] = useState<Compound[]>([]);
    const [templates, setTemplates] = useState<Template[]>([]);
    const [batchDataList, setBatchDataList] = useState<BatchData[]>([]);
    const [processingStatus, setProcessingStatus] = useState<string>('');

    // Fetch compounds on mount
    useEffect(() => {
        fetchCompounds();
    }, []);

    // Fetch templates when compound changes
    useEffect(() => {
        if (state.selectedCompound) {
            fetchTemplates(state.selectedCompound.id);
        }
    }, [state.selectedCompound]);

    const fetchCompounds = async () => {
        try {
            setState(prev => ({ ...prev, isLoading: true, error: undefined }));
            const response = await axios.get(`${API_BASE_URL}/api/compounds`);
            setCompounds(response.data.data || []);
        } catch (error) {
            setState(prev => ({ 
                ...prev, 
                error: 'Failed to fetch compounds. Please check your connection.' 
            }));
        } finally {
            setState(prev => ({ ...prev, isLoading: false }));
        }
    };

    const fetchTemplates = async (compoundId: string) => {
        try {
            setState(prev => ({ ...prev, isLoading: true, error: undefined }));
            const response = await axios.get(`${API_BASE_URL}/api/templates`, {
                params: { compound_id: compoundId }
            });
            setTemplates(response.data.data || []);
        } catch (error) {
            setState(prev => ({ 
                ...prev, 
                error: 'Failed to fetch templates.' 
            }));
        } finally {
            setState(prev => ({ ...prev, isLoading: false }));
        }
    };

    const handleCompoundSelect = (compound: Compound) => {
        setState(prev => ({ 
            ...prev, 
            selectedCompound: compound,
            selectedTemplate: undefined,
            extractedData: []
        }));
        setBatchDataList([]);
        setProcessingStatus('');
    };

    const handleTemplateSelect = (template: Template) => {
        setState(prev => ({ 
            ...prev, 
            selectedTemplate: template 
        }));
        setProcessingStatus('');
    };

    const handleProcessFiles = async () => {
        if (!state.selectedCompound || !state.selectedTemplate) {
            setState(prev => ({ 
                ...prev, 
                error: 'Please select compound and template first.' 
            }));
            return;
        }

        try {
            setState(prev => ({ ...prev, isLoading: true, error: undefined }));
            setProcessingStatus('Scanning PDF files and extracting batch analysis data...');
            
            // 启动批次分析处理
            const processResponse = await axios.post(
                `${API_BASE_URL}/api/documents/process-directory`,
                { 
                    compound_id: state.selectedCompound.id,
                    template_id: state.selectedTemplate.id
                }
            );

            const batchData = processResponse.data.data.batchData || [];
            setBatchDataList(batchData);
            
            setProcessingStatus(`Successfully processed ${batchData.length} batches!`);

        } catch (error) {
            setState(prev => ({ 
                ...prev, 
                error: 'Failed to process PDF files. Please try again.' 
            }));
            setProcessingStatus('');
        } finally {
            setState(prev => ({ ...prev, isLoading: false }));
        }
    };

    const handleInsertToWord = async () => {
        if (!state.selectedTemplate || batchDataList.length === 0) {
            setState(prev => ({ 
                ...prev, 
                error: 'No batch analysis data available to insert.' 
            }));
            return;
        }

        try {
            setState(prev => ({ ...prev, isLoading: true, error: undefined }));
            
            await Word.run(async (context) => {
                const selection = context.document.getSelection();
                
                // 生成基于模板的批次分析表格
                const tableHtml = generateBatchAnalysisTable(batchDataList);
                
                selection.insertHtml(tableHtml, Word.InsertLocation.replace);
                await context.sync();
            });
            
            setState(prev => ({ 
                ...prev, 
                error: undefined 
            }));

            setProcessingStatus('AIMTA Batch Analysis Tables inserted successfully!');
            
        } catch (error) {
            setState(prev => ({ 
                ...prev, 
                error: 'Failed to insert data into Word document.' 
            }));
        } finally {
            setState(prev => ({ ...prev, isLoading: false }));
        }
    };

    const generateBatchAnalysisTable = (batchData: BatchData[]): string => {
        if (!batchData.length) return '';

        // Table 1: Overview of Drug Substance Batches
        let table1Html = `
        <h2>Table 1: Overview of BGB-16673 Drug Substance Batches</h2>
        <table border="1" style="border-collapse: collapse; font-family: Arial, sans-serif; font-size: 10pt; width: 100%;">
            <thead>
                <tr style="background-color: #D9E1F2; font-weight: bold;">
                    <td style="padding: 8px; border: 1px solid #000;">Batch Number</td>
                    <td style="padding: 8px; border: 1px solid #000;">Batch Size (kg)</td>
                    <td style="padding: 8px; border: 1px solid #000;">Date of Manufacture</td>
                    <td style="padding: 8px; border: 1px solid #000;">Manufacturer</td>
                    <td style="padding: 8px; border: 1px solid #000;">Use(s)</td>
                </tr>
            </thead>
            <tbody>`;

        batchData.forEach(batch => {
            const manufacturerShort = batch.manufacturer.includes('Changzhou SynTheAll') ? 'Changzhou STA' : batch.manufacturer;
            table1Html += `
                <tr>
                    <td style="padding: 8px; border: 1px solid #000;">${batch.batch_number}</td>
                    <td style="padding: 8px; border: 1px solid #000;">TBD</td>
                    <td style="padding: 8px; border: 1px solid #000;">${batch.manufacture_date}</td>
                    <td style="padding: 8px; border: 1px solid #000;">${manufacturerShort}</td>
                    <td style="padding: 8px; border: 1px solid #000;">Clinical batch</td>
                </tr>`;
        });

        table1Html += `
            </tbody>
        </table>
        <p></p>`;

        // Table 2: Batch Analysis for GMP Batches
        let table2Html = `
        <h2>Table 2: Batch Analysis for GMP Batches of BGB-16673 Drug Substance</h2>
        <table border="1" style="border-collapse: collapse; font-family: Arial, sans-serif; font-size: 10pt; width: 100%;">
            <thead>
                <tr style="background-color: #D9E1F2; font-weight: bold;">
                    <td style="padding: 8px; border: 1px solid #000; width: 25%;">Test Parameter</td>
                    <td style="padding: 8px; border: 1px solid #000; width: 20%;">Acceptance Criterion</td>`;

        // Add batch number headers
        batchData.forEach(batch => {
            table2Html += `<td style="padding: 8px; border: 1px solid #000; text-align: center;">${batch.batch_number}</td>`;
        });

        table2Html += `
                </tr>
            </thead>
            <tbody>`;

        // Test parameters from template - Table 2
        const testParameters = [
            { name: "Appearance -- visual inspection", criterion: "Light yellow to yellow powder", key: "Appearance -- visual inspection" },
            { name: "Identification", criterion: "", key: "" },
            { name: "IR", criterion: "Conforms to reference standard", key: "IR" },
            { name: "HPLC", criterion: "Conforms to reference standard", key: "HPLC" },
            { name: "Assay -- HPLC (on anhydrous basis, %w/w)", criterion: "97.0-103.0", key: "Assay -- HPLC (on anhydrous basis, %w/w)" },
            { name: "Organic Impurities -- HPLC (%w/w)", criterion: "", key: "" },
            { name: "Single unspecified impurity", criterion: "≤ 0.50", key: "Single unspecified impurity" },
            { name: "BGB-24860", criterion: "", key: "BGB-24860" },
            { name: "RRT 0.56", criterion: "", key: "RRT 0.56" },
            { name: "RRT 0.70", criterion: "", key: "RRT 0.70" },
            { name: "RRT 0.72-0.73", criterion: "", key: "RRT 0.72-0.73" },
            { name: "RRT 0.76", criterion: "", key: "RRT 0.76" },
            { name: "RRT 0.80", criterion: "", key: "RRT 0.80" },
            { name: "RRT 1.10", criterion: "", key: "RRT 1.10" },
            { name: "Total impurities", criterion: "≤ 2.0", key: "Total impurities" },
            { name: "Enantiomeric Impurity -- HPLC (%w/w)", criterion: "≤ 1.0", key: "Enantiomeric Impurity -- HPLC (%w/w)" },
            { name: "Residual Solvents -- GC (ppm)", criterion: "", key: "" },
            { name: "Dichloromethane", criterion: "≤ 600", key: "Dichloromethane" },
            { name: "Ethyl acetate", criterion: "≤ 5000", key: "Ethyl acetate" },
            { name: "Isopropanol", criterion: "≤ 5000", key: "Isopropanol" },
            { name: "Methanol", criterion: "≤ 3000", key: "Methanol" },
            { name: "Tetrahydrofuran", criterion: "≤ 720", key: "Tetrahydrofuran" }
        ];

        testParameters.forEach(param => {
            table2Html += `
                <tr>
                    <td style="padding: 8px; border: 1px solid #000; font-weight: bold;">${param.name}</td>
                    <td style="padding: 8px; border: 1px solid #000;">${param.criterion}</td>`;

            batchData.forEach(batch => {
                const result = param.key ? (batch.test_results[param.key] || 'TBD') : '';
                table2Html += `<td style="padding: 8px; border: 1px solid #000;">${result}</td>`;
            });

            table2Html += `</tr>`;
        });

        table2Html += `
            </tbody>
        </table>
        <p></p>
        
        <h3>Table 2: Batch Analysis for GMP Batches of BGB-16673 Drug Substance (Continued)</h3>
        <table border="1" style="border-collapse: collapse; font-family: Arial, sans-serif; font-size: 10pt; width: 100%;">
            <thead>
                <tr style="background-color: #D9E1F2; font-weight: bold;">
                    <td style="padding: 8px; border: 1px solid #000; width: 25%;">Test Parameter</td>
                    <td style="padding: 8px; border: 1px solid #000; width: 20%;">Acceptance Criterion</td>`;

        // Add batch number headers for continued table
        batchData.forEach(batch => {
            table2Html += `<td style="padding: 8px; border: 1px solid #000; text-align: center;">${batch.batch_number}</td>`;
        });

        table2Html += `
                </tr>
            </thead>
            <tbody>`;

        // Table 2 Continued parameters
        const continuedParameters = [
            { name: "Residue on Ignition (%w/w)", criterion: "≤ 0.2", key: "Residue on Ignition (%w/w)" },
            { name: "Elemental Impurities -- ICP-MS", criterion: "", key: "" },
            { name: "Palladium (ppm)", criterion: "≤ 25", key: "Palladium (ppm)" },
            { name: "Polymorphic Form -- XRPD", criterion: "Conforms to reference standard", key: "Polymorphic Form -- XRPD" },
            { name: "Water Content -- KF (%w/w)", criterion: "Report result", key: "Water Content -- KF (%w/w)" }
        ];

        continuedParameters.forEach(param => {
            table2Html += `
                <tr>
                    <td style="padding: 8px; border: 1px solid #000; font-weight: bold;">${param.name}</td>
                    <td style="padding: 8px; border: 1px solid #000;">${param.criterion}</td>`;

            batchData.forEach(batch => {
                const result = param.key ? (batch.test_results[param.key] || 'TBD') : '';
                table2Html += `<td style="padding: 8px; border: 1px solid #000;">${result}</td>`;
            });

            table2Html += `</tr>`;
        });

        table2Html += `
            </tbody>
        </table>
        <p></p>`;

        // Add abbreviations for Table 2
        table2Html += `
        <p style="font-size: 9pt; font-style: italic;">
        <strong>Abbreviations:</strong> GC = gas chromatography; HPLC = high-performance liquid chromatography; 
        ICP-MS = inductively coupled plasma mass spectrometry; IR = infrared spectroscopy; KF = Karl Fischer; 
        ND = not detected; Pd = Palladium; RRT = relative retention time; XRPD = X‑ray powder diffraction.
        </p>
        <p></p>`;

        // Table 3: Batch Results for GMP Batches (Single batch format from template)
        let table3Html = `
        <h2>Table 3: Batch Results for GMP Batches of BGB-16673 Drug Substance</h2>`;

        // Generate Table 3 for the latest/most recent batch (or you can modify this logic)
        const latestBatch = batchData[batchData.length - 1]; // Take the last batch as example

        table3Html += `
        <table border="1" style="border-collapse: collapse; font-family: Arial, sans-serif; font-size: 10pt; width: 100%;">
            <thead>
                <tr style="background-color: #D9E1F2; font-weight: bold;">
                    <td style="padding: 8px; border: 1px solid #000; width: 40%;">Test Parameter</td>
                    <td style="padding: 8px; border: 1px solid #000; width: 30%;">Acceptance Criterion</td>
                    <td style="padding: 8px; border: 1px solid #000; width: 30%;">${latestBatch.batch_number}</td>
                </tr>
            </thead>
            <tbody>`;

        // Table 3 parameters (based on template structure)
        const table3Parameters = [
            { name: "Appearance -- visual inspection", criterion: "Light yellow to yellow powder", key: "Appearance -- visual inspection" },
            { name: "Identification", criterion: "", key: "" },
            { name: "IR", criterion: "Conforms to reference standard", key: "IR" },
            { name: "HPLC", criterion: "Conforms to reference standard", key: "HPLC" },
            { name: "Assay -- HPLC (on anhydrous basis, % w/w)", criterion: "97.0-103.0", key: "Assay -- HPLC (on anhydrous basis, %w/w)" },
            { name: "Organic Impurities -- HPLC (% w/w)", criterion: "", key: "" },
            { name: "Single unspecified impurity", criterion: "≤ 0.30", key: "Single unspecified impurity" },
            { name: "RRT 0.83", criterion: "", key: "RRT 0.83" },
            { name: "Total impurities", criterion: "≤ 2.0", key: "Total impurities" },
            { name: "Enantiomeric Impurity -- HPLC (% w/w)", criterion: "≤ 0.5", key: "Enantiomeric Impurity -- HPLC (%w/w)" },
            { name: "Residual Solvents -- GC (ppm)", criterion: "", key: "" },
            { name: "Dichloromethane", criterion: "≤ 600", key: "Dichloromethane" },
            { name: "Ethyl acetate", criterion: "≤ 5000", key: "Ethyl acetate" },
            { name: "Isopropanol", criterion: "≤ 5000", key: "Isopropanol" },
            { name: "Methanol", criterion: "≤ 3000", key: "Methanol" },
            { name: "Tetrahydrofuran", criterion: "≤ 720", key: "Tetrahydrofuran" },
            { name: "Inorganic Impurities", criterion: "", key: "" },
            { name: "Residue on ignition (% w/w)", criterion: "≤ 0.2", key: "Residue on Ignition (%w/w)" },
            { name: "Elemental impurities -- ICP-MS (Pd) (ppm)", criterion: "≤ 25", key: "Palladium (ppm)" },
            { name: "Polymorphic Form -- XRPD", criterion: "Conforms to reference standard", key: "Polymorphic Form -- XRPD" },
            { name: "Water Content -- KF (% w/w)", criterion: "≤ 3.5", key: "Water Content -- KF (%w/w)" }
        ];

        table3Parameters.forEach(param => {
            const result = param.key ? (latestBatch.test_results[param.key] || 'TBD') : '';
            table3Html += `
                <tr>
                    <td style="padding: 8px; border: 1px solid #000; font-weight: bold;">${param.name}</td>
                    <td style="padding: 8px; border: 1px solid #000;">${param.criterion}</td>
                    <td style="padding: 8px; border: 1px solid #000;">${result}</td>
                </tr>`;
        });

        table3Html += `
            </tbody>
        </table>
        <p></p>`;

        // Add abbreviations for Table 3
        const finalAbbreviations = `
        <p style="font-size: 9pt; font-style: italic;">
        <strong>Abbreviations:</strong> GC = gas chromatography; HPLC = high-performance liquid chromatography; 
        ICP-MS = inductively coupled plasma mass spectrometry; IR = infrared spectroscopy; KF = Karl Fischer; 
        ND = not detected; Pd = Palladium; RRT = relative retention time; XRPD = X‑ray powder diffraction.
        </p>`;

        return table1Html + table2Html + table3Html + finalAbbreviations;
    };

    return (
        <Stack className="app-container" tokens={{ childrenGap: 20 }}>
            <Stack.Item>
                <Text variant="xLarge" className="app-title">AIMTA Batch Analysis Processor</Text>
            </Stack.Item>

            {state.error && (
                <MessageBar 
                    messageBarType={MessageBarType.error}
                    onDismiss={() => setState(prev => ({ ...prev, error: undefined }))}
                >
                    {state.error}
                </MessageBar>
            )}

            {state.isLoading && (
                <Stack horizontalAlign="center" tokens={{ padding: 20 }}>
                    <Spinner size={SpinnerSize.large} label="Processing..." />
                </Stack>
            )}

            {processingStatus && (
                <MessageBar messageBarType={MessageBarType.info}>
                    {processingStatus}
                </MessageBar>
            )}

            <Stack tokens={{ childrenGap: 15 }}>
                <CompoundSelector
                    compounds={compounds}
                    selectedCompound={state.selectedCompound}
                    onSelect={handleCompoundSelect}
                    disabled={state.isLoading}
                />

                {state.selectedCompound && (
                    <TemplateSelector
                        templates={templates}
                        selectedTemplate={state.selectedTemplate}
                        onSelect={handleTemplateSelect}
                        disabled={state.isLoading}
                    />
                )}

                {state.selectedCompound && state.selectedTemplate && (
                    <Stack tokens={{ childrenGap: 10 }}>
                        <PrimaryButton
                            text="Process AIMTA Files"
                            iconProps={{ iconName: 'Processing' }}
                            onClick={handleProcessFiles}
                            disabled={state.isLoading}
                            styles={{ root: { width: '100%' } }}
                        />

                        {batchDataList.length > 0 && (
                            <Stack>
                                <MessageBar messageBarType={MessageBarType.success}>
                                    {`Ready to insert: ${batchDataList.length} batches analyzed`}
                                </MessageBar>
                                <PrimaryButton
                                    text="Insert Batch Analysis Tables"
                                    iconProps={{ iconName: 'Table' }}
                                    onClick={handleInsertToWord}
                                    disabled={state.isLoading}
                                    styles={{ root: { width: '100%' } }}
                                />
                            </Stack>
                        )}
                    </Stack>
                )}
            </Stack>

            <Stack className="help-text">
                <Text variant="small">• Select compound and template region</Text>
                <Text variant="small">• Click "Process AIMTA Files" to analyze batch data from PDFs</Text>
                <Text variant="small">• Each batch maintains independent data (no merging)</Text>
                <Text variant="small">• Click "Insert Batch Analysis Tables" to add complete tables to Word</Text>
                <Text variant="small">• Processing details shown in console</Text>
            </Stack>
        </Stack>
    );
};
