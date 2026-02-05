#!/usr/bin/env python3
"""
Feedback Analysis Tool

This script helps you analyze user feedback to improve the SOC-1 extraction process.

NOTE: This tool provides insights and suggestions, but does NOT automatically
modify your agent. You review the feedback and manually update prompts/code
in agent.py as needed. This gives you control and prevents unintended changes.

Usage:
    python analyze_feedback.py              # Show summary statistics
    python analyze_feedback.py --detailed   # Show detailed feedback entries
    python analyze_feedback.py --export     # Export to CSV for further analysis
"""

import json
import sys
from pathlib import Path
from collections import Counter
from datetime import datetime


def load_feedback():
    """Load feedback data from JSON log."""
    feedback_log = Path("feedback/feedback_log.json")
    
    if not feedback_log.exists():
        print("No feedback data found. Feedback log will be created when users submit feedback.")
        return []
    
    with open(feedback_log, "r") as f:
        return json.load(f)


def print_summary(feedback_data):
    """Print summary statistics."""
    if not feedback_data:
        print("No feedback entries yet.")
        return
    
    total = len(feedback_data)
    ratings = [f["rating"] for f in feedback_data]
    avg_rating = sum(ratings) / len(ratings)
    
    # Count issues
    all_issues = []
    for f in feedback_data:
        all_issues.extend(f.get("issues", []))
    issue_counts = Counter(all_issues)
    
    # Count feedback with corrected files
    corrected_count = sum(1 for f in feedback_data if f.get("corrected_file"))
    
    print("\n" + "="*60)
    print("FEEDBACK SUMMARY")
    print("="*60)
    print(f"\nTotal Feedback Entries: {total}")
    print(f"Average Rating: {avg_rating:.2f} / 5.0")
    print(f"Rating Distribution:")
    for rating in range(5, 0, -1):
        count = ratings.count(rating)
        percentage = (count / total) * 100
        bar = "█" * int(percentage / 2)
        print(f"  {rating} ★: {bar} {count} ({percentage:.1f}%)")
    
    print(f"\nCorrected Files Submitted: {corrected_count}")
    
    if issue_counts:
        print(f"\nMost Common Issues:")
        for issue, count in issue_counts.most_common():
            percentage = (count / total) * 100
            print(f"  • {issue.replace('_', ' ').title()}: {count} ({percentage:.1f}%)")
    
    # Recent trend
    if len(feedback_data) >= 5:
        recent_ratings = [f["rating"] for f in feedback_data[-5:]]
        recent_avg = sum(recent_ratings) / len(recent_ratings)
        print(f"\nRecent Trend (last 5):")
        print(f"  Average Rating: {recent_avg:.2f} / 5.0")
        
        if recent_avg > avg_rating:
            print("  📈 Improving!")
        elif recent_avg < avg_rating:
            print("  📉 Declining - needs attention")
        else:
            print("  ➡️  Stable")
    
    print("\n" + "="*60 + "\n")


def print_detailed(feedback_data):
    """Print detailed feedback entries."""
    if not feedback_data:
        print("No feedback entries yet.")
        return
    
    print("\n" + "="*60)
    print("DETAILED FEEDBACK")
    print("="*60 + "\n")
    
    for i, entry in enumerate(feedback_data, 1):
        timestamp = datetime.fromisoformat(entry["timestamp"])
        print(f"Entry #{i} - {timestamp.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"Job ID: {entry['job_id']}")
        print(f"Rating: {'★' * entry['rating']}{'☆' * (5 - entry['rating'])} ({entry['rating']}/5)")
        
        if entry.get("issues"):
            print(f"Issues: {', '.join(entry['issues'])}")
        
        if entry.get("feedback_text"):
            print(f"Comments: {entry['feedback_text']}")
        
        if entry.get("corrected_file"):
            print(f"Corrected File: ✓ Provided")
        
        metadata = entry.get("job_metadata", {})
        if metadata:
            print(f"Files: {metadata.get('type_ii_report')} → {metadata.get('management_review')}")
            
            summary = metadata.get("analysis_summary", {})
            if summary:
                print(f"Extraction Stats:")
                print(f"  - Controls: {summary.get('total_controls', 'N/A')}")
                print(f"  - Exceptions: {summary.get('exceptions', 'N/A')}")
                print(f"  - CUECs: {summary.get('total_cuecs', 'N/A')}")
        
        print("-" * 60 + "\n")


def export_to_csv(feedback_data):
    """Export feedback to CSV for analysis."""
    import csv
    
    if not feedback_data:
        print("No feedback to export.")
        return
    
    output_file = Path("feedback/feedback_export.csv")
    
    with open(output_file, "w", newline="") as f:
        writer = csv.writer(f)
        writer.writerow([
            "Timestamp", "Job ID", "Rating", "Issues", "Comments",
            "Has Corrected File", "Total Controls", "Exceptions", "CUECs"
        ])
        
        for entry in feedback_data:
            metadata = entry.get("job_metadata", {})
            summary = metadata.get("analysis_summary", {})
            
            writer.writerow([
                entry["timestamp"],
                entry["job_id"],
                entry["rating"],
                "; ".join(entry.get("issues", [])),
                entry.get("feedback_text", ""),
                "Yes" if entry.get("corrected_file") else "No",
                summary.get("total_controls", ""),
                summary.get("exceptions", ""),
                summary.get("total_cuecs", ""),
            ])
    
    print(f"\n✓ Feedback exported to: {output_file}")
    print(f"  {len(feedback_data)} entries exported\n")


def show_improvement_suggestions(feedback_data):
    """
    Analyze feedback and suggest improvements.
    
    NOTE: These are suggestions only. You should:
    1. Review the feedback and corrected files
    2. Manually update prompts in agent.py
    3. Test your changes
    4. Deploy when ready
    
    This keeps you in control and prevents automatic changes that could
    introduce bugs or degrade quality.
    """
    if not feedback_data:
        return
    
    print("\n" + "="*60)
    print("IMPROVEMENT SUGGESTIONS (Manual Review Required)")
    print("="*60 + "\n")
    
    # Analyze low ratings
    low_ratings = [f for f in feedback_data if f["rating"] <= 2]
    if low_ratings:
        print(f"⚠️  {len(low_ratings)} low ratings (≤2 stars) - Priority attention needed!")
        
        # Common issues in low ratings
        low_rating_issues = []
        for f in low_ratings:
            low_rating_issues.extend(f.get("issues", []))
        
        if low_rating_issues:
            issue_counts = Counter(low_rating_issues)
            print("\n   Most common issues in low ratings:")
            for issue, count in issue_counts.most_common(3):
                print(f"   • {issue.replace('_', ' ').title()}: {count} occurrences")
    
    # Analyze all issues
    all_issues = []
    for f in feedback_data:
        all_issues.extend(f.get("issues", []))
    
    if all_issues:
        issue_counts = Counter(all_issues)
        print("\n📊 Focus Areas (by frequency):")
        
        for issue, count in issue_counts.most_common():
            percentage = (count / len(feedback_data)) * 100
            
            if issue == "missing_controls":
                print(f"\n   1. Missing Controls ({percentage:.0f}% of feedback)")
                print("      → Review PDF extraction logic for control tables")
                print("      → Check if control numbering patterns are being missed")
                print("      → Consider adding fallback extraction methods")
            
            elif issue == "incorrect_mapping":
                print(f"\n   2. Incorrect Mapping ({percentage:.0f}% of feedback)")
                print("      → Review AI prompt for field mapping accuracy")
                print("      → Add more examples to the prompt")
                print("      → Consider adding validation rules")
            
            elif issue == "low_confidence":
                print(f"\n   3. Low Confidence Cells ({percentage:.0f}% of feedback)")
                print("      → Improve confidence scoring algorithm")
                print("      → Add more context to AI prompts")
                print("      → Consider human-in-the-loop for low confidence items")
            
            elif issue == "missing_cuecs":
                print(f"\n   4. Missing CUECs ({percentage:.0f}% of feedback)")
                print("      → Review CUEC extraction patterns")
                print("      → Check for alternative CUEC naming conventions")
                print("      → Add CUEC-specific extraction pass")
    
    # Check for corrected files
    corrected_files = [f for f in feedback_data if f.get("corrected_file")]
    if corrected_files:
        print(f"\n💡 {len(corrected_files)} corrected files available for training!")
        print("   → Use these to create few-shot examples for the AI")
        print("   → Analyze differences between original and corrected versions")
        print("   → Build a test suite from these examples")
    
    print("\n" + "="*60 + "\n")


def main():
    """Main entry point."""
    args = sys.argv[1:]
    
    feedback_data = load_feedback()
    
    if not feedback_data:
        print("\n📭 No feedback data available yet.")
        print("   Users will see a feedback form after downloading their results.")
        print("   Feedback will be stored in: feedback/feedback_log.json\n")
        return
    
    if "--detailed" in args:
        print_detailed(feedback_data)
    elif "--export" in args:
        export_to_csv(feedback_data)
    elif "--suggestions" in args or "--improve" in args:
        show_improvement_suggestions(feedback_data)
    else:
        print_summary(feedback_data)
        show_improvement_suggestions(feedback_data)


if __name__ == "__main__":
    main()
