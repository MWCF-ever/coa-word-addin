import React, { useState, useEffect } from 'react';
import { Stack, Text, MessageBar, MessageBarType, Spinner, SpinnerSize } from '@fluentui/react';
import { CompoundSelector } from './CompoundSelector';
import { TemplateSelector } from './TemplateSelector';
import { FileUploader } from './FileUploader';
import { ResultDisplay } from './ResultDisplay';
import { AppState, Compound, Template, ExtractedField, API_BASE_URL } from '../../types';
import axios from 'axios';
import '../taskpane.css';

export const App: React.FC = () => {
    const [state, setState] = useState<AppState>({
        extractedData: [],
        isLoading: false
    });

    const [compounds, setCompounds] = useState<Compound[]>([]);
    const [templates, setTemplates] = useState<Template[]>([]);

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
            uploadedDocument: undefined,
            extractedData: []
        }));
    };

    const handleTemplateSelect = (template: Template) => {
        setState(prev => ({ 
            ...prev, 
            selectedTemplate: template 
        }));
    };

    const handleFileUpload = async (file: File) => {
        if (!state.selectedCompound || !state.selectedTemplate) {
            setState(prev => ({ 
                ...prev, 
                error: 'Please select compound and template first.' 
            }));
            return;
        }

        try {
            setState(prev => ({ ...prev, isLoading: true, error: undefined }));
            
            const formData = new FormData();
            formData.append('file', file);
            formData.append('compound_id', state.selectedCompound.id);
            formData.append('template_id', state.selectedTemplate.id);

            // Upload file
            const uploadResponse = await axios.post(
                `${API_BASE_URL}/api/documents/upload`, 
                formData,
                {
                    headers: {
                        'Content-Type': 'multipart/form-data'
                    }
                }
            );

            const documentId = uploadResponse.data.data.documentId;

            // Process document
            const processResponse = await axios.post(
                `${API_BASE_URL}/api/documents/process`,
                { document_id: documentId }
            );

            setState(prev => ({ 
                ...prev, 
                uploadedDocument: uploadResponse.data.data,
                extractedData: processResponse.data.data.extractedData || []
            }));

        } catch (error) {
            setState(prev => ({ 
                ...prev, 
                error: 'Failed to process document. Please try again.' 
            }));
        } finally {
            setState(prev => ({ ...prev, isLoading: false }));
        }
    };

    const handleDataUpdate = (updatedData: ExtractedField[]) => {
        setState(prev => ({ ...prev, extractedData: updatedData }));
    };

    const handleInsertToWord = async () => {
        try {
            await Word.run(async (context) => {
                const selection = context.document.getSelection();
                
                // Insert template if selected
                if (state.selectedTemplate) {
                    selection.insertText(state.selectedTemplate.templateContent, Word.InsertLocation.replace);
                }
                
                // Insert extracted data
                state.extractedData.forEach(field => {
                    selection.insertText(`\n${field.fieldName}: ${field.fieldValue}`, Word.InsertLocation.end);
                });
                
                await context.sync();
            });
            
            setState(prev => ({ 
                ...prev, 
                error: undefined 
            }));
        } catch (error) {
            setState(prev => ({ 
                ...prev, 
                error: 'Failed to insert data into Word document.' 
            }));
        }
    };

    return (
        <Stack className="app-container" tokens={{ childrenGap: 20 }}>
            <Stack.Item>
                <Text variant="xLarge" className="app-title">COA Document Processor</Text>
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
                    <FileUploader
                        onUpload={handleFileUpload}
                        disabled={state.isLoading}
                    />
                )}

                {state.extractedData.length > 0 && (
                    <ResultDisplay
                        extractedData={state.extractedData}
                        onDataUpdate={handleDataUpdate}
                        onInsertToWord={handleInsertToWord}
                    />
                )}
            </Stack>
        </Stack>
    );
};