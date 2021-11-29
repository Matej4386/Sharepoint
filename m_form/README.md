## m-form

Webpart to render Sharepoint list forms (Display, Edit, New) for Sharepoint 2019 onpremise and Sharepoint online.
Supports:
inline editting in Display form,
2 rendering types: normal and inline (label and input are inline),
configurable via webpart properties with hooks (valueUpdated, beforeSave, ...),
it uses AddValidateUpdateItemUsingPath and ValidateUpdateListItem calls - columns conditions are supported (for example number columns must be between 10 and 100)

Supported fields are:
#### 'Text'
####  'Note' (Simple, RichText, AppendOnly - with previous text render)
#### 'User'
#### 'UserMulti'
#### 'Boolean'
#### 'DateTime' (Date + Date and Time)
#### 'Choice'
#### 'MultiChoice'
#### 'Number'
#### 'Currency'
#### 'Attachments' (Add, Edit, Delete)
#### 'Lookup':
#####    'TaxonomyFieldType'
#####    'TaxonomyFieldTypeMulti'
#####    'Lookup'
#####    'LookupMulti'
#### 'URL' (picture, link)

### Building the code

```bash
git clone the repo
npm i
gulp serve
```
