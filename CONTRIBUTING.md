# Contributing to vba-mcp-server

Thank you for your interest in contributing! / コントリビュートに興味を持っていただきありがとうございます！

## How to Contribute

### Reporting Bugs

1. Check if the issue already exists in [Issues](../../issues)
2. If not, create a new issue with:
   - Clear description of the problem
   - Steps to reproduce
   - Expected vs actual behavior
   - Your environment (Windows version, Office version, .NET version)

### Suggesting Features

1. Open a new issue with the "feature request" label
2. Describe the feature and its use case
3. Provide examples if possible

### Submitting Code

1. Fork the repository
2. Create a feature branch: `git checkout -b feature/your-feature-name`
3. Make your changes
4. Test your changes
5. Commit with clear messages
6. Push to your fork
7. Create a Pull Request

## Development Setup

### Prerequisites

- Windows 10/11
- .NET 8 SDK
- Visual Studio 2022 or VS Code
- Microsoft Office (for testing)

### Building

```bash
cd src/VbaMcpServer
dotnet build
```

### Testing

```bash
cd tests/VbaMcpServer.Tests
dotnet test
```

## Code Style

- Follow C# coding conventions
- Use meaningful variable and method names
- Add XML documentation comments for public APIs
- Keep methods focused and small

## Commit Messages

Use clear, descriptive commit messages:

```
feat: add procedure-level read/write support
fix: handle empty modules correctly  
docs: update installation instructions
refactor: simplify COM service error handling
```

## Questions?

Feel free to open an issue for any questions about contributing.
