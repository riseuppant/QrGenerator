document.addEventListener('DOMContentLoaded', () => {
  // DOM Element Selectors
  const elements = {
    addEventForm: document.getElementById('add-event-form'),
    qrForm: document.getElementById('qr-form'),
    downloadForm: document.getElementById('download-form'),
    eventMessage: document.getElementById('event-message'),
    generateMessage: document.getElementById('generate-message'),
    eventSelect: document.getElementById('event-select'),
    downloadSelect: document.getElementById('download-select')
  };

  // Utility Functions
  const showMessage = (messageElement, text, type) => {
    messageElement.textContent = text;
    messageElement.className = `message ${type}`;
  };

  const updateEventDropdowns = (events) => {
    const dropdowns = [elements.eventSelect, elements.downloadSelect];
    
    dropdowns.forEach(select => {
      // Remove all options except the first (default)
      while (select.options.length > 1) {
        select.remove(1);
      }

      // Add new events
      events.forEach(event => {
        const option = document.createElement('option');
        option.value = event;
        option.textContent = event;
        select.appendChild(option);
      });
    });
  };

  const handleServerError = (error, messageElement) => {
    console.error('Server error:', error);
    showMessage(messageElement, 'An unexpected error occurred', 'error');
  };

  // Event Addition Handler
  const handleEventAddition = async (e) => {
    e.preventDefault();
    const formData = new FormData(elements.addEventForm);
    
    showMessage(elements.eventMessage, 'Adding event...', 'loading');

    try {
      const response = await fetch('/add_event', {
        method: 'POST',
        body: formData
      });
      
      const result = await response.json();

      if (result.status === 'success' || result.status === 'warning') {
        // Update event dropdowns
        const eventSelects = [
          document.getElementById('event-select'), 
          document.getElementById('download-select')
        ];
        
        eventSelects.forEach(select => {
          // Clear existing options except the first
          while (select.options.length > 1) {
            select.remove(1);
          }
          
          // Add new events
          result.events.forEach(event => {
            const option = document.createElement('option');
            option.value = event;
            option.textContent = event;
            select.appendChild(option);
          });
        });

        // Display appropriate message
        if (result.status === 'warning') {
          let warningMessage = `Event added with ${result.total_rows - result.kept_rows} duplicate(s) removed.\n`;
          if (result.duplicates) {
            warningMessage += 'Duplicate entries:\n';
            result.duplicates.forEach(dup => {
              warningMessage += `Name: ${dup.Name}, Roll Number: ${dup['Roll Number']}\n`;
            });
          }
          showMessage(elements.eventMessage, warningMessage, 'warning');
        } else {
          showMessage(elements.eventMessage, result.message, 'success');
        }

        // Reset form
        elements.addEventForm.reset();
      } else {
        // Existing error handling
        let errorMessage = result.message;
        
        if (result.duplicates) {
          errorMessage = 'Duplicate roll numbers found:\n' + 
            result.duplicates.map(dup => 
              `Name: ${dup.Name}, Roll Number: ${dup['Roll Number']}`
            ).join('\n');
        }
        
        showMessage(elements.eventMessage, errorMessage, 'error');
      }
    } catch (error) {
      handleServerError(error, elements.eventMessage);
    }
  };

  // QR Code Generation Handler
  const handleQRGeneration = async (e) => {
    e.preventDefault();
    const formData = new FormData(elements.qrForm);
    
    showMessage(elements.generateMessage, 'Generating QR codes...', 'loading');

    try {
      const response = await fetch('/generate_qr', {
        method: 'POST',
        body: formData
      });
      
      const result = await response.json();

      showMessage(
        elements.generateMessage, 
        result.message, 
        result.status === 'success' ? 'success' : 'error'
      );
    } catch (error) {
      handleServerError(error, elements.generateMessage);
    }
  };

  // QR Code Download Handler
  const handleQRDownload = async (e) => {
    e.preventDefault();
    const selectedEvent = elements.downloadSelect.value;

    try {
      const response = await fetch(`/download_qr/${selectedEvent}`);
      
      if (response.ok) {
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        
        const downloadLink = document.createElement('a');
        downloadLink.href = url;
        downloadLink.download = `${selectedEvent}_qrcodes.zip`;
        
        document.body.appendChild(downloadLink);
        downloadLink.click();
        downloadLink.remove();
        
        window.URL.revokeObjectURL(url);
      } else {
        const errorData = await response.json();
        alert(errorData.message);
      }
    } catch (error) {
      console.error('Download error:', error);
      alert('Error downloading QR codes');
    }
  };

  // Event Listeners
  if (elements.addEventForm) {
    elements.addEventForm.addEventListener('submit', handleEventAddition);
  }

  if (elements.qrForm) {
    elements.qrForm.addEventListener('submit', handleQRGeneration);
  }

  if (elements.downloadForm) {
    elements.downloadForm.addEventListener('submit', handleQRDownload);
  }
});