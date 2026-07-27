from typing import Dict, List, Tuple
import pandas as pd


class DocumentComparator:
    """Compares documents between two Excel reports."""

    @staticmethod
    def parse_files(file_string: str) -> set:
        """
        Parse comma-separated file names into a set.
        
        Args:
            file_string: Comma-separated file names (e.g., "file1.pdf, file2.docx")
            
        Returns:
            Set of file names (stripped of whitespace)
        """
        if not file_string or pd.isna(file_string):
            return set()
        
        files = [f.strip() for f in str(file_string).split(',') if f.strip()]
        return set(files)

    @staticmethod
    def compare_documents(
        df_old: pd.DataFrame,
        df_new: pd.DataFrame
    ) -> Tuple[List[Dict], Dict]:
        """
        Compare two DataFrames and identify changes.
        Handles multiple rows per document name.
        
        Args:
            df_old: DataFrame from first report (earlier date)
            df_new: DataFrame from second report (later date)
            
        Returns:
            Tuple of (changes list, summary dict)
        """
        changes = []
        summary = {
            'total_changed': 0,
            'total_changes': 0,
            'revision_only': 0,
            'files_only': 0,
            'both_changed': 0
        }

        # Create dictionaries grouping by document name
        # Each document name maps to a list of (revision, files, approval_code) tuples
        old_dict = {}
        for idx, row in df_old.iterrows():
            doc_name = row['Document Name']
            revision = row['Revision']
            files = DocumentComparator.parse_files(row['File name'])
            approval_code = row.get('Approval Code', '(none)') if 'Approval Code' in row else '(none)'
            
            if doc_name not in old_dict:
                old_dict[doc_name] = []
            old_dict[doc_name].append({
                'revision': revision,
                'files': files,
                'approval_code': approval_code
            })

        new_dict = {}
        for idx, row in df_new.iterrows():
            doc_name = row['Document Name']
            revision = row['Revision']
            files = DocumentComparator.parse_files(row['File name'])
            approval_code = row.get('Approval Code', '(none)') if 'Approval Code' in row else '(none)'
            
            if doc_name not in new_dict:
                new_dict[doc_name] = []
            new_dict[doc_name].append({
                'revision': revision,
                'files': files,
                'approval_code': approval_code
            })

        # Sort by revision to handle zigzag patterns
        for doc_name in old_dict:
            old_dict[doc_name].sort(key=lambda x: (str(x['revision']), str(x['files'])))
        
        for doc_name in new_dict:
            new_dict[doc_name].sort(key=lambda x: (str(x['revision']), str(x['files'])))

        # Compare all documents from new report
        for doc_name, new_data_list in new_dict.items():
            if doc_name in old_dict:
                old_data_list = old_dict[doc_name]
                
                # Convert to sets for comparison
                old_set = set((d['revision'], frozenset(d['files'])) for d in old_data_list)
                new_set = set((d['revision'], frozenset(d['files'])) for d in new_data_list)
                
                # Check if there are any differences
                if old_set != new_set:
                    change = DocumentComparator._build_change_record_multiple(
                        doc_name,
                        old_data_list,
                        new_data_list
                    )
                    changes.append(change)
                    summary['total_changed'] += 1
                    summary['total_changes'] += 1
                    
                    # Determine change type
                    old_revisions = set(d['revision'] for d in old_data_list)
                    new_revisions = set(d['revision'] for d in new_data_list)
                    
                    if old_revisions != new_revisions:
                        summary['revision_only'] += 1
                    else:
                        summary['files_only'] += 1
            else:
                # New document added
                change = DocumentComparator._build_new_document_record_multiple(doc_name, new_data_list)
                changes.append(change)
                summary['total_changed'] += 1
                summary['total_changes'] += 1
                summary['files_only'] += 1

        # Check for removed documents
        for doc_name, old_data_list in old_dict.items():
            if doc_name not in new_dict:
                change = DocumentComparator._build_removed_document_record_multiple(doc_name, old_data_list)
                changes.append(change)
                summary['total_changed'] += 1
                summary['total_changes'] += 1
                summary['files_only'] += 1

        return changes, summary

    @staticmethod
    def _build_change_record_multiple(
        doc_name: str,
        old_data_list: List[Dict],
        new_data_list: List[Dict]
    ) -> Dict:
        """
        Build a change record for a modified document with multiple revisions.
        
        Args:
            doc_name: Document name
            old_data_list: List of old document data (multiple revisions)
            new_data_list: List of new document data (multiple revisions)
            
        Returns:
            Change record dictionary
        """
        # Get all revisions and files
        old_revisions = sorted(set(d['revision'] for d in old_data_list))
        new_revisions = sorted(set(d['revision'] for d in new_data_list))
        
        # Get approval codes
        old_approval_codes = sorted(set(d['approval_code'] for d in old_data_list if d['approval_code'] != '(none)'))
        new_approval_codes = sorted(set(d['approval_code'] for d in new_data_list if d['approval_code'] != '(none)'))
        
        old_approval_code_str = ', '.join(old_approval_codes) if old_approval_codes else "(none)"
        new_approval_code_str = ', '.join(new_approval_codes) if new_approval_codes else "(none)"
        
        # Combine all files
        old_all_files = set()
        for d in old_data_list:
            old_all_files.update(d['files'])
        
        new_all_files = set()
        for d in new_data_list:
            new_all_files.update(d['files'])
        
        old_files_str = ', '.join(sorted(old_all_files)) if old_all_files else "(none)"
        new_files_str = ', '.join(sorted(new_all_files)) if new_all_files else "(none)"
        
        # Calculate file changes
        files_added = new_all_files - old_all_files
        files_removed = old_all_files - new_all_files
        
        # Build remarks
        remarks_parts = []
        
        if old_revisions != new_revisions:
            old_rev_str = ', '.join(str(r) for r in old_revisions)
            new_rev_str = ', '.join(str(r) for r in new_revisions)
            remarks_parts.append(f"Revisions changed from [{old_rev_str}] to [{new_rev_str}]")
        
        if files_added:
            remarks_parts.append(f"Files added: {', '.join(sorted(files_added))}")
        
        if files_removed:
            remarks_parts.append(f"Files removed: {', '.join(sorted(files_removed))}")
        
        remarks = "; ".join(remarks_parts)
        
        # Determine change type
        revision_changed = old_revisions != new_revisions
        files_changed = old_all_files != new_all_files
        
        if revision_changed and files_changed:
            change_type = "BOTH_CHANGED"
        elif revision_changed:
            change_type = "REVISION_CHANGED"
        else:
            change_type = "FILES_CHANGED"
        
        return {
            'document_name': doc_name,
            'old_approval_code': old_approval_code_str,
            'new_approval_code': new_approval_code_str,
            'old_revision': ', '.join(str(r) for r in old_revisions),
            'new_revision': ', '.join(str(r) for r in new_revisions),
            'old_files': old_files_str,
            'new_files': new_files_str,
            'change_type': change_type,
            'remarks': remarks
        }

    @staticmethod
    def _build_new_document_record_multiple(doc_name: str, new_data_list: List[Dict]) -> Dict:
        """
        Build a record for a newly added document with multiple revisions.
        
        Args:
            doc_name: Document name
            new_data_list: List of new document data (multiple revisions)
            
        Returns:
            Change record dictionary
        """
        new_revisions = sorted(set(d['revision'] for d in new_data_list))
        
        # Get approval codes
        new_approval_codes = sorted(set(d['approval_code'] for d in new_data_list if d['approval_code'] != '(none)'))
        new_approval_code_str = ', '.join(new_approval_codes) if new_approval_codes else "(none)"
        
        new_all_files = set()
        for d in new_data_list:
            new_all_files.update(d['files'])
        
        new_files_str = ', '.join(sorted(new_all_files)) if new_all_files else "(none)"
        
        return {
            'document_name': doc_name,
            'old_approval_code': "(new)",
            'new_approval_code': new_approval_code_str,
            'old_revision': "(new)",
            'new_revision': ', '.join(str(r) for r in new_revisions),
            'old_files': "(none)",
            'new_files': new_files_str,
            'change_type': "FILES_CHANGED",
            'remarks': f"New document added with revisions [{', '.join(str(r) for r in new_revisions)}]"
        }

    @staticmethod
    def _build_removed_document_record_multiple(doc_name: str, old_data_list: List[Dict]) -> Dict:
        """
        Build a record for a removed document with multiple revisions.
        
        Args:
            doc_name: Document name
            old_data_list: List of old document data (multiple revisions)
            
        Returns:
            Change record dictionary
        """
        old_revisions = sorted(set(d['revision'] for d in old_data_list))
        
        # Get approval codes
        old_approval_codes = sorted(set(d['approval_code'] for d in old_data_list if d['approval_code'] != '(none)'))
        old_approval_code_str = ', '.join(old_approval_codes) if old_approval_codes else "(none)"
        
        old_all_files = set()
        for d in old_data_list:
            old_all_files.update(d['files'])
        
        old_files_str = ', '.join(sorted(old_all_files)) if old_all_files else "(none)"
        
        return {
            'document_name': doc_name,
            'old_approval_code': old_approval_code_str,
            'new_approval_code': "(removed)",
            'old_revision': ', '.join(str(r) for r in old_revisions),
            'new_revision': "(removed)",
            'old_files': old_files_str,
            'new_files': "(none)",
            'change_type': "FILES_CHANGED",
            'remarks': "Document removed"
        }
