import React, { useState } from 'react';
import { 
    Stack, 
    Text, 
    TextField, 
    PrimaryButton, 
    DefaultButton,
    ProgressIndicator,
    Icon,
    Separator,
    MessageBar,
    MessageBarType
} from '@fluentui/react';
import { ExtractedField, API_BASE_URL } from '../../types';
import axios from 'axios';

interface ResultDisplayProps {
    extractedData: ExtractedField[];
    onDataUpdate: (updatedData: ExtractedField[]) => void;
    onInsertToWord: () => void;
    documentId?: string;
}

export const ResultDisplay: React.FC<ResultDisplayProps> = ({
    extractedData,
    onDataUpdate,
    onInsertToWord,
    documentId
}) => {
    const [editedData, setEditedData] = useState<Record<string, string>>({});
    const [isInserting, setIsInserting] = useState(false);
    const [insertMessage, setInsertMessage] = useState<string | null>(null);
    const [insertError, setInsertError] = useState<string | null>(null);

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

    const handleInsertToWord = async () => {
        console.log('开始插入Word表格...');
        console.log('Document ID:', documentId);
        
        if (!documentId) {
            setInsertError('Document ID not available');
            console.error('Document ID 缺失');
            return;
        }

        setIsInserting(true);
        setInsertError(null);
        setInsertMessage(null);

        // 在父作用域声明变量
        let insertLocation = '';
        let insertSuccess = false;

        try {
            // 1. 从后端获取表格数据
            console.log('正在获取表格数据...');
            const apiUrl = `${API_BASE_URL}/api/documents/${documentId}/word-table-data`;
            console.log('API URL:', apiUrl);
            
            const response = await axios.get(apiUrl, {
                timeout: 30000,
                headers: {
                    'Accept': 'application/json',
                    'Content-Type': 'application/json'
                }
            });
            
            console.log('API响应状态:', response.status);
            console.log('API响应数据:', response.data);
            
            if (!response.data.success) {
                throw new Error(response.data.error || 'API returned failure');
            }
            
            const tableData = response.data.data.tableData;
            console.log('表格数据:', tableData);
            
            if (!tableData || !Array.isArray(tableData)) {
                throw new Error('Invalid table data format');
            }

            // 2. 使用HTML方法插入到Word文档
            console.log('正在使用HTML方法插入Word文档...');
            await Word.run(async (context) => {
                try {
                    // 获取当前选择位置
                    const range = context.document.getSelection();
                    console.log('获取Word选择范围成功');
                    
                    // 创建HTML表格字符串
                    let htmlTable = `
                    <table border="1" style="border-collapse: collapse; font-family: Arial, sans-serif; font-size: 10pt; width: 100%;">
                        <thead>
                            <tr style="background-color: #D9E1F2; font-weight: bold;">
                                <td style="padding: 8px; border: 1px solid #000; width: 40%;">Field / 字段</td>
                                <td style="padding: 8px; border: 1px solid #000; width: 60%;">Value / 值</td>
                            </tr>
                        </thead>
                        <tbody>
                    `;
                    
                    // 添加数据行
                    for (const item of tableData) {
                        let cellValue = item.value || '';
                        if (item.confidence !== undefined && item.confidence > 0) {
                            cellValue += ` (${(item.confidence * 100).toFixed(0)}%)`;
                        }
                        
                        // 根据置信度设置背景颜色
                        let bgColor = '';
                        if (item.confidence !== undefined) {
                            if (item.confidence < 0.6) {
                                bgColor = 'background-color: #FFE6E6;'; // 浅红色
                            } else if (item.confidence < 0.8) {
                                bgColor = 'background-color: #FFF4E6;'; // 浅橙色
                            }
                        }
                        
                        htmlTable += `
                            <tr>
                                <td style="padding: 8px; border: 1px solid #000; font-weight: bold;">${item.field || ''}</td>
                                <td style="padding: 8px; border: 1px solid #000; ${bgColor}">${cellValue}</td>
                            </tr>
                        `;
                    }
                    
                    htmlTable += `
                        </tbody>
                    </table>
                    <p></p>
                    `;
                    
                    // 在当前位置插入HTML表格
                    range.insertHtml(htmlTable, Word.InsertLocation.after);
                    await context.sync();
                    
                    insertSuccess = true;
                    insertLocation = '当前选择位置';
                    console.log('HTML表格插入成功');
                    
                } catch (wordError: any) {
                    console.error('Word操作错误:', wordError);
                    throw new Error(`Word操作失败: ${wordError?.message || wordError}`);
                }
            });

            setInsertMessage(`COA表格已成功插入到Word文档 (${insertLocation || '未知位置'})！`);
            console.log('表格插入完成！');
            
        } catch (error: any) {
            console.error('插入Word表格失败:', error);
            
            let errorMessage = '插入表格到Word文档失败，请重试。';
            
            if (axios.isAxiosError(error)) {
                console.error('Axios错误详情:', {
                    message: error.message,
                    code: error.code,
                    status: error.response?.status,
                    statusText: error.response?.statusText,
                    data: error.response?.data
                });
                
                if (error.response?.status === 404) {
                    errorMessage = '文档未找到，请重新上传并处理文档。';
                } else if (error.response?.status === 500) {
                    errorMessage = `服务器错误: ${error.response?.data?.detail || '内部错误'}`;
                } else if (error.code === 'ECONNABORTED') {
                    errorMessage = '请求超时，请检查网络连接后重试。';
                } else {
                    errorMessage = `网络错误 (${error.response?.status || 'Unknown'}): ${error.message}`;
                }
            } else if (error instanceof Error) {
                errorMessage = error.message;
            }
            
            setInsertError(errorMessage);
        } finally {
            setIsInserting(false);
        }
    };

    const getFieldDisplayName = (fieldName: string): string => {
        const displayNames: Record<string, string> = {
            'lot_number': 'Lot Number / 批号',
            'manufacturer': 'Manufacturer / 生产商',
            'storage_condition': 'Storage Condition / 储存条件',
            'Manufacture_Date': 'Manufacture Date / 生产日期'
        };
        return displayNames[fieldName] || fieldName;
    };

    const getConfidenceColor = (score: number): string => {
        if (score >= 0.9) return '#107C10';
        if (score >= 0.7) return '#FFA500';
        return '#D13438';
    };

    // 检查是否有编辑过的数据
    const hasUnsavedChanges = Object.keys(editedData).length > 0;

    return (
        <Stack tokens={{ childrenGap: 15 }}>
            <Separator />
            <Text variant="large" className="section-title">Extracted Data</Text>
            
            {/* 调试信息 */}
            <MessageBar messageBarType={MessageBarType.info}>
                Document ID: {documentId || 'Not available'}
            </MessageBar>
            
            {/* 显示插入结果消息 */}
            {insertMessage && (
                <MessageBar 
                    messageBarType={MessageBarType.success}
                    onDismiss={() => setInsertMessage(null)}
                >
                    {insertMessage}
                </MessageBar>
            )}
            
            {insertError && (
                <MessageBar 
                    messageBarType={MessageBarType.error}
                    onDismiss={() => setInsertError(null)}
                >
                    {insertError}
                </MessageBar>
            )}
            
            {/* 进度指示器 */}
            {isInserting && (
                <ProgressIndicator label="正在插入表格到Word文档..." />
            )}
            
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
                    text={isInserting ? "Inserting..." : "Insert Table to Word"}
                    iconProps={{ iconName: 'Table' }}
                    onClick={handleInsertToWord}
                    disabled={isInserting || !documentId}
                    styles={{ root: { flex: 1 } }}
                />
                <DefaultButton
                    text="Save Changes"
                    iconProps={{ iconName: 'Save' }}
                    onClick={handleSaveChanges}
                    disabled={!hasUnsavedChanges || isInserting}
                    styles={{ root: { flex: 1 } }}
                />
            </Stack>

            <Stack className="help-text">
                <Text variant="small">
                    • 表格将插入到Word文档的当前光标位置
                </Text>
                <Text variant="small">
                    • 请确保在Word文档中选择合适的插入位置
                </Text>
                <Text variant="small">
                    • 低置信度字段用黄色/红色背景标记，请仔细检查
                </Text>
                <Text variant="small">
                    • 点击"Save Changes"保存修改的字段值
                </Text>
            </Stack>
        </Stack>
    );
};
