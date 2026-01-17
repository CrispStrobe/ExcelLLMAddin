# LLM Excel Add-in

Connect Excel to multiple LLM providers (Ollama, OpenAI, Mistral, etc.) using VBA.

## Features

- ✅ Works on Mac (probably also on Windows, to be tested...)
- ✅ Multiple LLM providers (Ollama, OpenAI, Mistral, Nebius, Scaleway, OpenRouter)
- ✅ Easy configuration via menu system
- ✅ Custom functions: `=PROMPT()`, `=LIST_MODELS()`, `=LLM_CONFIG()`
- ✅ Full debug logging

## Requirements

### For Ollama (Free & Local)
1. Install Ollama: https://ollama.com
2. Install a model: `ollama pull llama3.2`
3. Start Ollama: `ollama serve`

### For Cloud Providers (OpenAI, etc.)
- Get API key from provider
- Configure in Settings menu

### Mac Requirements
- curl (pre-installed on macOS)
- Excel for Mac 2016 or later

### Windows Requirements
- curl (built-in or download from https://curl.se)
- Excel 2016 or later

## Installation

### Mac
1. Download `ExcelLLMAddin.xlam`
2. Double-click to open it in Excel
3. Go to **Tools > Add-Ins**
4. Check **LLM Excel Add-in** to enable
5. Save location: `~/Library/Group Containers/UBF8T346G9.Office/User Content/Add-Ins/`

### Windows
1. Download `ExcelLLMAddin.xlam`
2. Open Excel
3. Go to **File > Options > Add-Ins**
4. Click **Go** next to "Excel Add-ins"
5. Click **Browse** and select `ExcelLLMAddin.xlam`
6. Check the box to enable it

## Configuration

Run from **Tools > Macro > Macros**:

1. **ShowSettings** - Configure providers and models
2. **QuickTest** - Test your connection
3. **TestCurlConnection** - Verify curl is working

## Usage

### In Cells
```excel
=PROMPT("What is 5+5?")
=PROMPT("Translate 'Hello' to German", "ollama", "llama3.2")
=PROMPT(A1)  ' Use cell reference
```

### Keyboard Shortcuts (may not work on all Macs)
- ⌘+Shift+L = Settings
- ⌘+Shift+T = Quick Test

## Functions

### `PROMPT(text, [provider], [model])`
Send a prompt to the LLM and get a response.

**Examples:**
```excel
=PROMPT("Explain quantum physics")
=PROMPT("Summarize this: " & A1)
=PROMPT(A1, "ollama", "llama3.2")
```

### `LIST_MODELS([provider])`
List available models from a provider.

**Example:**
```excel
=LIST_MODELS("ollama")
```

### `LLM_CONFIG()`
Show current provider and model.

**Example:**
```excel
=LLM_CONFIG()
```

## Troubleshooting

### "Error: Cannot create HTTP object"
- **Mac**: Update the code to use curl (already included)
- **Windows**: Make sure curl is installed

### "Error: Request timeout"
- Make sure Ollama is running: `ollama serve`
- Check if model is installed: `ollama list`
- Increase timeout in code (currently 60s)

### "Model not found"
- List your models: `ollama list`
- Pull the model: `ollama pull llama3.2`
- Update model in Settings (Tools > Macro > ShowSettings)

### German Umlaute not working
- Use the updated version with CleanEncodingIssues function
- System message now tells LLM to avoid emojis

### Debug Mode
The add-in has extensive debug logging. To view:
1. Press **⌘+G** (Mac) or **Ctrl+G** (Windows) in VBA
2. View the Immediate Window for detailed logs

## Configuration File

Settings are stored in:
- **Mac**: `~/Library/Containers/com.microsoft.Excel/Data/ExcelLLMAddin_config.txt`
- **Windows**: `%USERPROFILE%\ExcelLLMAddin_config.txt`

You can edit this file manually or use ShowSettings.

## Supported Providers

| Provider | Free? | Local? | API Key Required? |
|----------|-------|--------|-------------------|
| Ollama | ✅ | ✅ | ❌ |
| OpenAI | ❌ | ❌ | ✅ |
| Mistral | ❌ | ❌ | ✅ |
| Nebius | ❌ | ❌ | ✅ |
| Scaleway | ❌ | ❌ | ✅ |
| OpenRouter | ❌ | ❌ | ✅ |

## License

MIT License - feel free to modify and distribute!
