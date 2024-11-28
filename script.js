document.getElementById("add-chord").addEventListener("click", () => {
    const note = document.getElementById("note").value;
    const modifier = document.getElementById("modifier").value;

    // Format the chord as "NoteModifier"
    const chord = `${note}${modifier}`;
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            const selectedText = result.value;
            Office.context.document.setSelectedDataAsync(
                `${chord}\n${selectedText}`,
                { coercionType: Office.CoercionType.Text }
            );
        } else {
            console.error(result.error.message);
        }
    });
});
