# OCLC Work ID Retrieval Script

The scripts in this project use various user-selected files (.xlsx and .mrc) to query OCLC's WorldCat Search API v.2 for the Work ID associated with each title in the file. 

## Description
Addressing title matches across editions and formats presents an ongoing challenge for library technical service staff. Whether synchronizing a bookstore's adoption list with the library's catalog or simplifying the reconciliation of electronic and physical holdings, identifying a reliable data point for comparing editions and material types can be daunting. This project's scripts utilize OCLC's WorldCat Search API v.2, conducting queries to pinpoint bibliographic records with matching ISBNs and extract the corresponding "Work ID" from the identified record.

## Getting Started

### Dependencies

* The WorldCat Search API v.2 requires a client ID and secret that are only available to libraries that maintain an OCLC Cataloging and Metadata subscription (full cataloging) and a FirstSearch/WorldCat Discovery subscription. Both are required.
* Requirements for the OAuth 2.0 portions of these scripts can be found on the [OCLC Developer Network](https://github.com/OCLC-Developer-Network/gists/tree/22a524aff103cc42b3e9be0b600fc13143e6e157/authentication/python) github page.
* Scripts require the creation of a "config.yml" file consisting of key, secret, auth_url, token_url, and metadata_service_url. A sample yaml configuration file can be found on the [OCLC Developer Network](https://github.com/OCLC-Developer-Network/gists/tree/22a524aff103cc42b3e9be0b600fc13143e6e157/authentication) github page.

## License

This project is licensed under the Apache-2.0 License - see the LICENSE.md file for details

## Acknowledgments

* OCLC OAuth 2.0 code belongs in its entirety to [Karen Coombs](https://github.com/librarywebchic) and has NOT be substantially changed from its original form.
* All other idea generation, workflow execution, and code production belongs to [Jason C. Mitchell](https://github.com/CompareTheo)
