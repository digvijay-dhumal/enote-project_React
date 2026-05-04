
export const fetchNoteData = async (noteId) => {
    const response = await fetch('https://indianbankxenciaapp.azurewebsites.net/api/ENote/getENoteGeneralDetails', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify({ "noteId": noteId }),
    });
    const data = await response.json();
    return data;
};

