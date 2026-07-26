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

        # Create dictionaries for quick lookup
        old_dict = {}
        for idx, row in df_old.iterrows():
            doc_name = row['Document Name']
            old_dict[doc_name] = {
                'revision': row['Revision'],
                'files': DocumentComparator.parse_files(row['File name'])
            }

        new_dict = {}
        for idx, row in df_new.iterrows():
            doc_name = row['Document Name']
            new_dict[doc_name] = {
                'revision': row['Revision'],
                'files': DocumentComparator.parse_files(row['File name'])
            }

        # Compare all documents from new report
        for doc_name, new_data in new_dict.items():
            if doc_name in old_dict:
                old_data = old_dict[doc_name]
                
                # Check for changes
                revision_changed = old_data['revision'] != new_data['revision']
                files_changed = old_data['files'] != new_data['files']

                if revision_changed or files_changed:
                    change = DocumentComparator._build_change_record(
                        doc_name,
                        old_data,
                        new_data,
                        revision_changed,
                        files_changed
                    )
                    changes.append(change)
                    summary['total_changed'] += 1
                    summary['total_changes'] += 1

                    if revision_changed and files_changed:
                        summary['both_changed'] += 1
                    elif revision_changed:
                        summary['revision_only'] += 1
                    else:
                        summary['files_only'] += 1
            else:
                # New document added (files changed, revision is new)
                change = DocumentComparator._build_new_document_record(doc_name, new_data)
                changes.append(change)
                summary['total_changed'] += 1
                summary['total_changes'] += 1
                summary['files_only'] += 1

        # Check for removed documents
        for doc_name, old_data in old_dict.items():
            if doc_name not in new_dict:
                change = DocumentComparator._build_removed_document_record(doc_name, old_data)
                changes.append(change)
                summary['total_changed'] += 1
                summary['total_changes'] += 1
                summary['files_only'] += 1

        return changes, summary

    @staticmethod
    def _build_change_record(
        doc_name: str,
        old_data: Dict,
        new_data: Dict,
        revision_changed: bool,
        files_changed: bool
    ) -> Dict:
        """
        Build a change record for a modified document.
        
        Args:
            doc_name: Document name
            old_data: Old document data
            new_data: New document data
            revision_changed: Whether revision changed
            files_changed: Whether files changed
            
        Returns:
            Change record dictionary
        """
        old_files_str = ', '.join(sorted(old_data['files'])) if old_data['files'] else "(none)"
        new_files_str = ', '.join(sorted(new_data['files'])) if new_data['files'] else "(none)"

        # Calculate file changes
        files_added = new_data['files'] - old_data['files']
        files_removed = old_data['files'] - new_data['files']

        # Build remarks
        remarks_parts = []
        
        if revision_changed:
            remarks_parts.append(
                f"Revision changed from {old_data['revision']} to {new_data['revision']}"
            )
        
        if files_added:
            remarks_parts.append(f"Files added: {', '.join(sorted(files_added))}")
        
        if files_removed:
            remarks_parts.append(f"Files removed: {', '.join(sorted(files_removed))}")

        remarks = "; ".join(remarks_parts)

        # Determine change type
        if revision_changed and files_changed:
            change_type = "BOTH_CHANGED"
        elif revision_changed:
            change_type = "REVISION_CHANGED"
        else:
            change_type = "FILES_CHANGED"

        return {
            'document_name': doc_name,
            'old_revision': old_data['revision'],
            'new_revision': new_data['revision'],
            'old_files': old_files_str,
            'new_files': new_files_str,
            'change_type': change_type,
            'remarks': remarks
        }

    @staticmethod
    def _build_new_document_record(doc_name: str, new_data: Dict) -> Dict:
        """
        Build a record for a newly added document.
        
        Args:
            doc_name: Document name
            new_data: New document data
            
        Returns:
            Change record dictionary
        """
        new_files_str = ', '.join(sorted(new_data['files'])) if new_data['files'] else "(none)"

        return {
            'document_name': doc_name,
            'old_revision': "(new)",
            'new_revision': new_data['revision'],
            'old_files': "(none)",
            'new_files': new_files_str,
            'change_type': "FILES_CHANGED",
            'remarks': f"New document added with revision {new_data['revision']}"
        }

    @staticmethod
    def _build_removed_document_record(doc_name: str, old_data: Dict) -> Dict:
        """
        Build a record for a removed document.
        
        Args:
            doc_name: Document name
            old_data: Old document data
            
        Returns:
            Change record dictionary
        """
        old_files_str = ', '.join(sorted(old_data['files'])) if old_data['files'] else "(none)"

        return {
            'document_name': doc_name,
            'old_revision': old_data['revision'],
            'new_revision': "(removed)",
            'old_files': old_files_str,
            'new_files': "(none)",
            'change_type': "FILES_CHANGED",
            'remarks': "Document removed"
        }
