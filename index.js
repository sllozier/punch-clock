document.addEventListener('DOMContentLoaded', () => {
    const punchButton = document.getElementById('punchButton');
    const clearButton = document.getElementById('clearButton');
    const inRecordList = document.getElementById('inRecordList');
    const outRecordList = document.getElementById('outRecordList');
    const totalHoursElem = document.getElementById('totalHours');
    const exportButton = document.getElementById('exportButton');
    let records = JSON.parse(localStorage.getItem('records')) || [];
    let currentState = records.length % 2 === 0 ? 'in' : 'out';

    const saveRecords = () => {
        localStorage.setItem('records', JSON.stringify(records));
    };

    const formatDateTime = (dateString) => {
        const options = { weekday: 'long', year: 'numeric', month: 'numeric', day: 'numeric', hour: 'numeric', minute: 'numeric', hour12: true };
        return new Date(dateString).toLocaleString('en-US', options);
    };

    const formatDate = (dateString) => {
        const options = { year: 'numeric', month: 'numeric', day: 'numeric' };
        return new Date(dateString).toLocaleDateString('en-US', options);
    };

    const formatTime = (dateString) => {
        const options = { hour: 'numeric', minute: 'numeric', hour12: true };
        return new Date(dateString).toLocaleTimeString('en-US', options);
    };

    const formatDay = (dateString) => {
        const options = { weekday: 'long' };
        return new Date(dateString).toLocaleDateString('en-US', options);
    };

    const renderRecords = () => {
        inRecordList.innerHTML = '';
        outRecordList.innerHTML = '';
        records.forEach((record, index) => {
            const li = document.createElement('li');
            li.textContent = `${record.state.toUpperCase()} - ${formatDateTime(record.time)}`;
            li.setAttribute('data-index', index);
            li.setAttribute('contenteditable', false);

            li.addEventListener('click', () => {
                li.setAttribute('contenteditable', true);
                li.focus();
            });

            li.addEventListener('blur', () => {
                li.setAttribute('contenteditable', false);
                const newValue = li.textContent.split(' - ')[1];
                if (newValue) {
                    records[index].time = new Date(newValue).toISOString();
                    saveRecords();
                    renderRecords();
                    calculateTotalHours();
                }
            });

            li.addEventListener('keypress', (event) => {
                if (event.key === 'Enter') {
                    event.preventDefault();
                    li.blur();
                }
            });
            const deleteButton = document.createElement('button');
            deleteButton.textContent = 'x';
            deleteButton.addEventListener('click', () => {
                records.splice(index, 1);
                saveRecords();
                renderRecords();
                calculateTotalHours();
            });
            li.appendChild(deleteButton);
            if (record.state === 'in') {
                inRecordList.appendChild(li);
            } else {
                outRecordList.appendChild(li);
            }
        });
    };

    const calculateTotalHours = () => {
        let totalHours = 0;
        for (let i = 0; i < records.length; i += 2) {
            if (records[i + 1]) {
                const inTime = new Date(records[i].time);
                const outTime = new Date(records[i + 1].time);
                totalHours += (outTime - inTime) / (1000 * 60 * 60);
            }
        }
        totalHoursElem.textContent = `Total Hours: ${totalHours.toFixed(2)}`;
        return totalHours.toFixed(2);
    };

    const exportToExcel = () => {
        const wb = XLSX.utils.book_new();
        const ws_data = [
            ["IN", "", "", "OUT", "", "", ""],
            ["Day of the week", "Date", "Time", "Day of the week", "Date", "Time"]
        ];

        for (let i = 0; i < records.length; i += 2) {
            const row = [];
            if (records[i]) {
                row.push(formatDay(records[i].time));
                row.push(formatDate(records[i].time));
                row.push(formatTime(records[i].time));
            } else {
                row.push("", "", "");
            }
            if (records[i + 1]) {
                row.push(formatDay(records[i + 1].time));
                row.push(formatDate(records[i + 1].time));
                row.push(formatTime(records[i + 1].time));
            } else {
                row.push("", "", "");
            }
            ws_data.push(row);
        }

        ws_data.push(["Total Time:", calculateTotalHours()]);

        const ws = XLSX.utils.aoa_to_sheet(ws_data);
        ws['!merges'] = [
            { s: { r: 0, c: 0 }, e: { r: 0, c: 2 } }, // Merge "IN" across A, B, C
            { s: { r: 0, c: 3 }, e: { r: 0, c: 6 } }  // Merge "OUT" across D, E, F
        ];
        
        XLSX.utils.book_append_sheet(wb, ws, "Records");
        XLSX.writeFile(wb, "time_punch_records.xlsx");
    };

    punchButton.addEventListener('click', () => {
        const now = new Date().toISOString();
        if (records.length === 0 || currentState === 'in') {
            records.push({ state: 'in', time: now });
            currentState = 'out';
        } else {
            records.push({ state: 'out', time: now });
            currentState = 'in';
        }
        saveRecords();
        renderRecords();
        calculateTotalHours();
    });

    clearButton.addEventListener('click', () => {
        records = [];
        saveRecords();
        renderRecords();
        calculateTotalHours();
    });

    exportButton.addEventListener('click', exportToExcel);

    renderRecords();
    calculateTotalHours();

    // Analog Clock Functionality
    const secondHand = document.querySelector('.second-hand');
    const minuteHand = document.querySelector('.minute-hand');
    const hourHand = document.querySelector('.hour-hand');

    function setDate() {
        const now = new Date();
        
        const seconds = now.getSeconds();
        const secondsDegrees = ((seconds / 60) * 360) + 90; // Add 90 to offset for CSS transform
        secondHand.style.transform = `rotate(${secondsDegrees}deg)`;
        
        const minutes = now.getMinutes();
        const minutesDegrees = ((minutes / 60) * 360) + ((seconds / 60) * 6) + 90;
        minuteHand.style.transform = `rotate(${minutesDegrees}deg)`;
        
        const hours = now.getHours();
        const hoursDegrees = ((hours / 12) * 360) + ((minutes / 60) * 30) + 90;
        hourHand.style.transform = `rotate(${hoursDegrees}deg)`;
    }

    setInterval(setDate, 1000);
    setDate(); // Initial call to set the time immediately
});
