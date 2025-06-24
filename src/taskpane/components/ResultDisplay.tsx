import React, { useState } from 'react';
import { 
    Stack, 
    Text, 
    TextField, 
    PrimaryButton, 
    DefaultButton,
    ProgressIndicator,
    Icon,
    Separator
} from '@fluentui/react';
import { ExtractedField } from '../../types';

interface ResultDisplayProps {
    extractedData: ExtractedField[];
    onDataUpdate: (updatedData: ExtractedField[]) => void;
    onInsertToWord: () => void;
}

export const ResultDisplay: React.FC<ResultDisplayProps> = ({
    extractedData,
    onDataUpdate,
    onInsertToWord
}) => {
    const [editedData, setEditedData] = useState<Record<string, string>>({});

    const handleFieldChange = (fieldName: string, value: string) => {
        setEditedData(prev => ({
            ...prev,
            [fieldName]: value
        }));
    };

    const handleSaveChanges = () => {
        const updatedData = extractedData.map(field => ({
            ...field,
            fieldValue: editedData[field.fieldName] || field.fieldValue
        }));
        onDataUpdate(updatedData);
        setEditedData({});
    };

    const getFieldDisplayName = (fieldName: string): string => {
        const displayNames: Record<string, string> = {
            'lot_number': 'Lot Number / 批号',
            'manufacturer': 'Manufacturer / 生产商',
            'storage_condition': 'Storage Condition / 储存条件'
        };
        return displayNames[fieldName] || fieldName;
    };

    const getConfidenceColor = (score: number): string => {
        if (score >= 0.9) return '#107C10';
        if (score >= 0.7) return '#FFA500';
        return '#D13438';
    };

    return (
        <Stack tokens={{ childrenGap: 15 }}>
            <Separator />
            <Text variant="large" className="section-title">Extracted Data</Text>
            
            <Stack tokens={{ childrenGap: 15 }}>
                {extractedData.map((field) => (
                    <Stack key={field.id} tokens={{ childrenGap: 5 }}>
                        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                            <Text variant="medium">{getFieldDisplayName(field.fieldName)}</Text>
                            <Stack horizontal tokens={{ childrenGap: 5 }} verticalAlign="center">
                                <Icon 
                                    iconName="Info" 
                                    style={{ 
                                        color: getConfidenceColor(field.confidenceScore),
                                        fontSize: 12
                                    }} 
                                />
                                <Text 
                                    variant="small" 
                                    style={{ color: getConfidenceColor(field.confidenceScore) }}
                                >
                                    {(field.confidenceScore * 100).toFixed(0)}% confidence
                                </Text>
                            </Stack>
                        </Stack>
                        
                        <TextField
                            value={editedData[field.fieldName] || field.fieldValue}
                            onChange={(e, newValue) => handleFieldChange(field.fieldName, newValue || '')}
                            multiline={field.fieldName === 'storage_condition'}
                            rows={field.fieldName === 'storage_condition' ? 3 : 1}
                            styles={{
                                fieldGroup: {
                                    backgroundColor: field.confidenceScore < 0.7 ? '#FFF4E6' : undefined
                                }
                            }}
                        />
                        
                        {field.originalText && (
                            <Text variant="small" style={{ color: '#605E5C', fontStyle: 'italic' }}>
                                Original: "{field.originalText}"
                            </Text>
                        )}
                    </Stack>
                ))}
            </Stack>

            <Separator />

            <Stack horizontal tokens={{ childrenGap: 10 }}>
                <PrimaryButton
                    text="Insert to Word"
                    iconProps={{ iconName: 'WordDocument' }}
                    onClick={onInsertToWord}
                    styles={{ root: { flex: 1 } }}
                />
                <DefaultButton
                    text="Save Changes"
                    iconProps={{ iconName: 'Save' }}
                    onClick={handleSaveChanges}
                    disabled={Object.keys(editedData).length === 0}
                    styles={{ root: { flex: 1 } }}
                />
            </Stack>

            <Stack className="help-text">
                <Text variant="small">
                    • Fields highlighted in yellow have low confidence scores
                </Text>
                <Text variant="small">
                    • Click "Save Changes" to update the extracted values
                </Text>
                <Text variant="small">
                    • Click "Insert to Word" to add the data to your document
                </Text>
            </Stack>
        </Stack>
    );
};