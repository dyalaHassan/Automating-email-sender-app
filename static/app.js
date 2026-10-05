function showFileName(input, targetId) {

    const target = document.getElementById(targetId);

    if (!target) {
        return;
    }

    if (input.files && input.files.length > 0) {

        target.textContent =
            input.files[0].name;

    } else {

        target.textContent =
            "لم يتم اختيار ملف";
    }
}


function copyField(text, button) {

    navigator.clipboard
        .writeText(text)
        .then(() => {

            const message =
                document.getElementById(
                    "copy-message"
                );

            if (message) {

                message.textContent =
                    "✓ تم نسخ " + text;

                setTimeout(() => {

                    message.textContent = "";

                }, 1800);
            }
        });
}


function updateSendOptions() {

    const selected =
        document.querySelector(
            'input[name="send_type"]:checked'
        );

    const firstOptions =
        document.getElementById(
            "first-options"
        );

    const rangeOptions =
        document.getElementById(
            "range-options"
        );

    if (!selected) {
        return;
    }

    if (firstOptions) {
        firstOptions.classList.add(
            "hidden"
        );
    }

    if (rangeOptions) {
        rangeOptions.classList.add(
            "hidden"
        );
    }

    if (
        selected.value === "first" &&
        firstOptions
    ) {

        firstOptions.classList.remove(
            "hidden"
        );
    }

    if (
        selected.value === "range" &&
        rangeOptions
    ) {

        rangeOptions.classList.remove(
            "hidden"
        );
    }
}


document.addEventListener(
    "DOMContentLoaded",
    function () {

        updateSendOptions();

    }
);