module.exports = [
  {
    files: ['src/**/*.{js,gs}'],
    languageOptions: {
      ecmaVersion: 2020,
      sourceType: 'script',
      globals: {
        SpreadsheetApp: 'readonly',
        GmailApp: 'readonly',
        HtmlService: 'readonly',
        ContentService: 'readonly',
        UrlFetchApp: 'readonly',
        Logger: 'readonly',
        PropertiesService: 'readonly',
        DriveApp: 'readonly',
        Utilities: 'readonly',
        ScriptApp: 'readonly',
        Session: 'readonly',
        headerMap: 'readonly'
      }
    },
    rules: {
      'no-undef': 'error',
      'no-unused-vars': ['warn', { argsIgnorePattern: '^_' }],
      'eqeqeq': ['warn', 'always']
    }
  }
];
