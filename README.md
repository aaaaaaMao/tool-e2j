# tool-e2j

[reference](https://docs.sheetjs.com/docs/solutions/output#api-1)

## Install

```bash
npm install -g .
```

## Usage

### excel to json

#### basic

```bash
tool-e2j xx.xlsx xx.csv
```
- output format:
```json
[
  {
    "sheet_name": "Sheet1",
    "data": [
      {
        "a": 1,
        "b": 2
      }
    ]
  }
]
```

#### split sheet

```bash
tool-e2j --split xx.xlsx
```
- output format:
```json
[
  {
    "a": 1,
    "b": 2
  }
]
```
- output file name format: `{output file name}-${sheet name}.json`

### json to excel

#### basic

```bash
tool-e2j xx.json
```

- input format:
```json
[{
    "a": 1,
    "b": 2
}]
```

```json
{
    "a": 1,
    "b": 2
}
```

#### multi sheet

```bash
tool-e2j --multi xx.json
```

- input format:
```json
[
  {
    "sheet_name": "Sheet1",
    "data": [
      {
        "a": 1,
        "b": 2
      }
    ]
  }
]
```