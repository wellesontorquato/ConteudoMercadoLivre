window.addEventListener('DOMContentLoaded', () => {
  const fileInput = document.getElementById('file');
  const fileLabel = document.querySelector('.file-input-label');
  const fileName = document.querySelector('.file-name');
  const deleteButton = document.getElementById('deleteButton');
  const submitButton = document.querySelector('.form-submit');
  const loadingOverlay = document.querySelector('.loading-overlay');

  // Oculta o botão "Baixar planilha" ao carregar a página
  submitButton.style.display = 'none';

  fileInput.addEventListener('change', () => {
    const file = fileInput.files[0];
    if (file) {
      fileName.textContent = file.name;
      fileLabel.style.display = 'none';
      deleteButton.style.display = 'inline-block';
      fileName.style.border = '1px solid #FFBE10';

      // Mostra a animação de carregamento por 2 segundos
      loadingOverlay.style.opacity = '1';
      loadingOverlay.style.visibility = 'visible';
      submitButton.style.display = 'none';
      fileName.textContent = '';
      fileLabel.style.display = 'none';
      deleteButton.style.display = 'none';
      fileName.style.border = 'none';
      setTimeout(() => {
        loadingOverlay.style.opacity = '0';
        loadingOverlay.style.visibility = 'hidden';
        submitButton.style.display = 'block';
        fileName.textContent = file.name;
        fileLabel.style.display = 'none';
        deleteButton.style.display = 'inline-block';
        fileName.style.border = '1px solid #FFBE10';
      }, 2000);
    } else {
      fileName.textContent = '';
      fileLabel.style.display = 'block';
      deleteButton.style.display = 'none';
      fileName.style.border = 'none';
      submitButton.style.display = 'none';
    }
  });

  deleteButton.addEventListener('click', (event) => {
    event.preventDefault(); // Impede o envio do formulário
    fileInput.value = '';
    fileName.textContent = '';
    fileLabel.style.display = 'inline-block';
    deleteButton.style.display = 'none';
    fileLabel.style.width = '80%';
    fileName.style.border = 'none';
    submitButton.style.display = 'none';
  });

  // Limpa o arquivo selecionado quando o ícone do botão de exclusão é clicado
  deleteButton.addEventListener('click', () => {
    fileInput.value = '';
  });
});
