# Infinity.GetDeps

GetDeps is a wrapper for SysInternal's Process Monitor (procmon) for the use in the Windows PreInstallation Environment.

GetDeps will watch the list of events in a hidden procmon window and read those entries. On startup GetDeps will load a predefined configuration (pmc) for dropping filtered events, and including only event with the result NAME_NOT_FOUND on both registry API calls and file related calls. In short, allowing a deveoper to resolve system dependencies when extending the functionality of WinPE with ease.

InfinityResearchAndDevelopment 2017
