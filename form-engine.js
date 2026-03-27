// Chef SaaS Motion Tracker — Dynamic Form Engine

let _answers = {};
let _fields = [];
let _onAnswersChange = null;

// Special token evaluation for conditional visibility
function _shouldShow(field, currentAnswers) {
  if (!field.show_when_field) return true;

  const controlValue = (currentAnswers[field.show_when_field] || "").trim();
  const token = (field.show_when_value || "").trim();

  switch (token) {
    case "PITCHED_OR_WON":
      return controlValue.startsWith("Pitched") || controlValue === "Closed Won";
    case "BLOCKER_STATUSES":
      return (
        controlValue === "Pitched - No Interest" ||
        controlValue === "Pitched - Exploring" ||
        controlValue === "Closed Lost"
      );
    case "NOT_NOT_PITCHED":
      return controlValue !== "" && controlValue !== "Not Pitched";
    case "NEXT_STEP_STATUSES":
      return (
        controlValue === "Pitched - Exploring" ||
        controlValue === "Pitched - Active Deal"
      );
    default:
      return controlValue === token;
  }
}

function _updateVisibility(container) {
  _fields.forEach((field) => {
    if (!field.show_when_field) return;
    const wrapper = container.querySelector(`[data-field="${field.field_name}"]`);
    if (!wrapper) return;
    const visible = _shouldShow(field, _answers);
    wrapper.style.display = visible ? "block" : "none";
    if (!visible) {
      // Clear value when hidden so it doesn't pollute submission
      _answers[field.field_name] = "";
      const el = wrapper.querySelector("input, select, textarea");
      if (el) el.value = "";
    }
  });
}

function _buildInput(field, options) {
  const { field_name, field_type, is_required } = field;

  if (field_type === "dropdown") {
    const select = document.createElement("select");
    select.id = `field-${field_name}`;
    select.name = field_name;

    const placeholder = document.createElement("option");
    placeholder.value = "";
    placeholder.textContent = "— Select —";
    select.appendChild(placeholder);

    (options || []).forEach((opt) => {
      const option = document.createElement("option");
      option.value = opt.option_value;
      option.textContent = opt.option_value;
      select.appendChild(option);
    });

    return select;
  }

  if (field_type === "date") {
    const input = document.createElement("input");
    input.type = "date";
    input.id = `field-${field_name}`;
    input.name = field_name;
    return input;
  }

  if (field_type === "text_area") {
    const textarea = document.createElement("textarea");
    textarea.id = `field-${field_name}`;
    textarea.name = field_name;
    textarea.rows = 4;
    textarea.placeholder =
      "Type your notes here, or tap the 🎙️ mic on your keyboard to dictate";
    return textarea;
  }

  // Default: text input
  const input = document.createElement("input");
  input.type = "text";
  input.id = `field-${field_name}`;
  input.name = field_name;
  return input;
}

async function renderForm(container, playId, onAnswersChange) {
  _answers = {};
  _fields = [];
  _onAnswersChange = onAnswersChange;
  container.innerHTML = '<div class="spinner"></div>';

  try {
    _fields = await getFormFields(playId);

    // Pre-fetch all dropdown options in parallel
    const optionMap = {};
    await Promise.all(
      _fields
        .filter((f) => f.field_type === "dropdown")
        .map(async (f) => {
          optionMap[f.field_name] = await getFieldOptions(playId, f.field_name);
        })
    );

    container.innerHTML = "";

    _fields.forEach((field) => {
      const wrapper = document.createElement("div");
      wrapper.className = "form-field";
      wrapper.dataset.field = field.field_name;

      const label = document.createElement("label");
      label.htmlFor = `field-${field.field_name}`;
      label.innerHTML = field.field_label;
      if (field.is_required) {
        label.innerHTML += ' <span class="required-star">*</span>';
      }

      const input = _buildInput(field, optionMap[field.field_name]);

      input.addEventListener("change", () => {
        _answers[field.field_name] = input.value;
        _updateVisibility(container);
        if (_onAnswersChange) _onAnswersChange({ ..._answers });
      });

      input.addEventListener("input", () => {
        _answers[field.field_name] = input.value;
        if (_onAnswersChange) _onAnswersChange({ ..._answers });
      });

      wrapper.appendChild(label);
      wrapper.appendChild(input);
      container.appendChild(wrapper);

      // Apply initial visibility
      if (field.show_when_field) wrapper.style.display = "none";
    });
  } catch (err) {
    container.innerHTML = `<p class="error-msg">Failed to load form: ${err.message}</p>`;
  }
}

function getAnswers() {
  return { ..._answers };
}

function validateForm() {
  const errors = [];

  _fields.forEach((field) => {
    if (!field.is_required) return;

    // Only validate visible fields
    const wrapper = document.querySelector(`[data-field="${field.field_name}"]`);
    if (wrapper && wrapper.style.display === "none") return;

    const value = (_answers[field.field_name] || "").trim();
    if (!value) {
      errors.push(`"${field.field_label}" is required.`);
    }
  });

  return { valid: errors.length === 0, errors };
}
