import * as React from 'react';
import { useState, useEffect, useCallback, useRef } from 'react';
import { 
  Button, 
  Textarea, 
  Card, 
  Text, 
  ProgressBar,
  Badge,
  Divider,
  useToastController
} from '@fluentui/react-components';
import { 
  ArrowDownload24Regular, 
  ArrowUpload24Regular,
  Edit24Regular,
  Info24Regular
} from '@fluentui/react-icons';
import { useMermaidSelection } from '../hooks/useMermaidSelection';
import { insertMermaidDiagram, replaceMermaidDiagram } from '../mermaid';

const MermaidEditor = () => {
  const [code, setCode] = useState('');
  const [isGenerating, setIsGenerating] = useState(false);
  const [selectedDiagram, setSelectedDiagram] = useState(null);
  const [isUpdateMode, setIsUpdateMode] = useState(false);
  const [status, setStatus] = useState({ type: 'info', message: 'Enter Mermaid code to generate a diagram' });
  const { dispatchToast } = useToastController();
  
  // Debounce timer for code changes
  const debounceTimerRef = useRef(null);
  
  // Use the selection hook
  const { isLoading: selectionLoading, error: selectionError, processSelection } = useMermaidSelection(
    // Callback when diagram is detected
    useCallback((metadata, picture) => {
      setSelectedDiagram({ ...metadata, picture });
      setCode(metadata.code);
      setIsUpdateMode(true);
      setStatus({ 
        type: 'success', 
        message: 'Mermaid diagram detected. You can edit and update it.' 
      });
    }, []),
    // Callback when selection is cleared
    useCallback(() => {
      setSelectedDiagram(null);
      setIsUpdateMode(false);
      setStatus({ 
        type: 'info', 
        message: 'Enter Mermaid code to generate a new diagram' 
      });
    }, [])
  );

  /**
   * Handle code changes with debouncing
   */
  const handleCodeChange = useCallback((newCode) => {
    setCode(newCode);
    
    // Clear existing debounce timer
    if (debounceTimerRef.current) {
      clearTimeout(debounceTimerRef.current);
    }
    
    // Set new debounce timer for preview update
    debounceTimerRef.current = setTimeout(() => {
      // Preview would be updated here in a real implementation
      // For now, we just validate the code
      if (newCode.trim()) {
        setStatus({ type: 'info', message: 'Code updated. Ready to generate diagram.' });
      } else {
        setStatus({ type: 'info', message: 'Enter Mermaid code to generate a diagram' });
      }
    }, 500);
  }, []);

  /**
   * Generate or update diagram
   */
  const handleGenerateDiagram = useCallback(async () => {
    if (!code.trim()) {
      setStatus({ type: 'error', message: 'Please enter Mermaid code first.' });
      return;
    }

    setIsGenerating(true);
    setStatus({ type: 'info', message: 'Generating diagram...' });

    try {
      if (isUpdateMode && selectedDiagram) {
        // Update existing diagram
        await replaceMermaidDiagram(selectedDiagram.picture, code);
        setStatus({ type: 'success', message: 'Diagram updated successfully!' });
        dispatchToast({
          intent: 'success',
          title: 'Diagram Updated',
          message: 'The selected diagram has been replaced with the new version.'
        });
      } else {
        // Insert new diagram
        await insertMermaidDiagram(code);
        setStatus({ type: 'success', message: 'Diagram inserted successfully!' });
        dispatchToast({
          intent: 'success', 
          title: 'Diagram Inserted',
          message: 'The diagram has been added to your document.'
        });
      }
    } catch (error) {
      console.error('Failed to generate diagram:', error);
      setStatus({ 
        type: 'error', 
        message: `Failed to ${isUpdateMode ? 'update' : 'insert'} diagram: ${error.message}` 
      });
      dispatchToast({
        intent: 'error',
        title: 'Generation Failed',
        message: error.message || 'An error occurred while generating the diagram.'
      });
    } finally {
      setIsGenerating(false);
    }
  }, [code, isUpdateMode, selectedDiagram, dispatchToast]);

  /**
   * Handle selection errors
   */
  useEffect(() => {
    if (selectionError) {
      setStatus({ 
        type: 'error', 
        message: `Selection monitoring error: ${selectionError}` 
      });
    }
  }, [selectionError]);

  /**
   * Cleanup debounce timer on unmount
   */
  useEffect(() => {
    return () => {
      if (debounceTimerRef.current) {
        clearTimeout(debounceTimerRef.current);
      }
    };
  }, []);

  return (
    <div style={{ padding: '16px', display: 'flex', flexDirection: 'column', gap: '16px' }}>
      <Card>
        <div style={{ padding: '16px' }}>
          <Text size="large" weight="semibold">
            Mermaid Diagram Editor
          </Text>
          
          {isUpdateMode && (
            <div style={{ marginTop: '8px' }}>
              <Badge appearance="tint" color="green" size="small">
                <Info24Regular />
                Update Mode - Selected diagram detected
              </Badge>
            </div>
          )}
        </div>
      </Card>

      {selectionLoading && (
        <Card>
          <div style={{ padding: '16px' }}>
            <Text>Analyzing selection...</Text>
            <ProgressBar />
          </div>
        </Card>
      )}

      <Card>
        <div style={{ padding: '16px' }}>
          <Text weight="semibold">Mermaid Code</Text>
          <Textarea
            value={code}
            onChange={(e) => handleCodeChange(e.target.value)}
            placeholder="Enter your Mermaid diagram code here..."
            style={{ 
              marginTop: '8px', 
              minHeight: '200px', 
              fontFamily: 'monospace',
              fontSize: '14px'
            }}
            resize="vertical"
          />
        </div>
      </Card>

      {status.message && (
        <Card appearance="outline">
          <div style={{ padding: '16px' }}>
            <Text 
              weight="medium" 
              style={{ 
                color: status.type === 'error' ? '#d13438' : 
                       status.type === 'success' ? '#107c10' : '#0078d4' 
              }}
            >
              {status.message}
            </Text>
          </div>
        </Card>
      )}

      <div style={{ display: 'flex', gap: '8px', justifyContent: 'flex-end' }}>
        <Button
          appearance="secondary"
          onClick={() => processSelection()}
          disabled={selectionLoading}
        >
          Refresh Selection
        </Button>
        
        <Button
          appearance="primary"
          icon={isUpdateMode ? <Edit24Regular /> : <ArrowDownload24Regular />}
          onClick={handleGenerateDiagram}
          disabled={isGenerating || !code.trim() || selectionLoading}
        >
          {isGenerating ? 'Generating...' : (isUpdateMode ? 'Update Diagram' : 'Insert Diagram')}
        </Button>
      </div>

      <Divider />

      <Card appearance="outline">
        <div style={{ padding: '16px' }}>
          <Text weight="semibold">Quick Start</Text>
          <div style={{ marginTop: '8px', fontSize: '14px', lineHeight: '1.5' }}>
            <Text block>
              1. Enter Mermaid diagram code in the editor above<br/>
              2. Click "Insert Diagram" to add it to your document<br/>
              3. Select an existing diagram to edit and update it<br/>
              4. Diagrams are stored as images with metadata for editing
            </Text>
          </div>
        </div>
      </Card>
    </div>
  );
};

export default MermaidEditor;
