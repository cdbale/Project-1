"""Package the verified guide, official materials, and complete grading skill."""
import argparse
import hashlib
import re
from pathlib import Path, PurePosixPath
from zipfile import ZIP_DEFLATED, ZipFile

parser = argparse.ArgumentParser()
parser.add_argument('--skill-dir', required=True, type=Path)
args = parser.parse_args()
grading_dir = Path(__file__).resolve().parents[1]
project_dir = grading_dir.parent
skill_dir = args.skill_dir.resolve()
archive_path = grading_dir / 'Project_1_Grading_Resources.zip'
root = PurePosixPath('Project_1_Grading_Resources')
sources = {
    root / 'Project_1_TA_Grading_Guide.docx': grading_dir / 'Project_1_TA_Grading_Guide.docx',
}
for name in ('Project 1 Instructions.pdf', 'Rubric V2.pdf', 'ifood.csv',
             'Executive Summary Example.pdf', 'Technical_Report_Example.qmd'):
    sources[root / 'Materials' / name] = project_dir / name

required_skill_files = {
    'SKILL.md', 'agents/openai.yaml', 'scripts/score.py',
    'references/rubric-v2.md', 'references/project-context.md',
    'references/ta-workflow.md', 'references/canvas-intake-and-outputs.md',
    'references/overseer-and-assessment-agents.md',
}
packaged_skill_files = set()
for source in sorted(skill_dir.rglob('*')):
    if not source.is_file():
        continue
    relative = source.relative_to(skill_dir)
    if any(part.startswith('.') or part == '__pycache__' for part in relative.parts):
        continue
    if source.suffix in {'.pyc', '.pyo'} or source.name.endswith('~'):
        continue
    sources[root / 'skills' / skill_dir.name / relative.as_posix()] = source
    packaged_skill_files.add(relative.as_posix())
assert required_skill_files <= packaged_skill_files, 'Incomplete skill source'
assert all(source.is_file() for source in sources.values()), 'Missing package source'

# Confirm that the unchanged PDFs and data match the baseline bundled in the skill.
context = (skill_dir / 'references/project-context.md').read_text(encoding='utf-8')
for name in ('Project 1 Instructions.pdf', 'Rubric V2.pdf',
             'Executive Summary Example.pdf', 'ifood.csv'):
    match = re.search(r'\| `' + re.escape(name) + r'` \| `([a-f0-9]{64})` \|', context)
    assert match, f'Missing source fingerprint: {name}'
    assert hashlib.sha256((project_dir / name).read_bytes()).hexdigest() == match.group(1), name

readme = '''MKTG 411 Project 1 grading resources
Package prepared October 12, 2026

START HERE
Extract the entire ZIP into a private course folder on your computer.
Open Project_1_TA_Grading_Guide.docx and complete its Initial setup steps.
Keep the extracted Materials and complete skills folder structure intact.
Do not work directly inside the ZIP.

CONTENTS
Project_1_TA_Grading_Guide.docx: role-specific setup and grading workflow.
Materials/Project 1 Instructions.pdf: official assignment and data dictionary.
Materials/Rubric V2.pdf: official rubric.
Materials/ifood.csv: required raw starting data.
Materials/Executive Summary Example.pdf and Technical_Report_Example.qmd:
  teaching examples, not answer keys or additional grading requirements.
  The QMD and raw CSV are together so its relative data path is preserved.
skills/grade-mktg-411-project-1/: complete overseer skill, all references,
  agent metadata, and the deterministic Python score calculator.

AI SETUP
Give the AI the extracted folder path and ask it to read
skills/grade-mktg-411-project-1/SKILL.md and its linked references.
If the skill is not installed, replace its $grade-mktg-411-project-1
invocation in the guide's prompts with "use the skill at [full SKILL.md path]."
An installed copy must match this package's fresh-agent overseer version.
Use an environment that can create a fresh assessment agent per student
without inherited conversation history. If unsupported, assessment stays
pending as described in the skill; do not simulate isolated agents.

STORAGE AND SEPARATELY PROVIDED ITEMS
Let the AI ask where to create a separate Project_1_Grading folder for
student originals, AI_Assessment.xlsx files, assessment_result.json files,
and the root AI_Assigned_Scores.csv. Keep initial TA judgments outside
both the resource folder and the student grading folder, unavailable to AI.
Obtain Canvas access, the live roster, the previous semester's Google Drive
spreadsheet, and any new instructor clarifications separately. Credentials
and student records are not included. The TA team copies the previous sheet
for this semester, preserving median formulas. TAs personally reconcile
their grades, enter final TA totals, and post instructor-approved medians.

The penguin sentence is a review indicator only and changes no points.
Each student's final project grade is the median of final TA total grades.
AI scores do not receive an extra vote.
'''

with ZipFile(archive_path, 'w', ZIP_DEFLATED, compresslevel=9) as archive:
    for target, source in sorted(sources.items(), key=lambda item: str(item[0])):
        archive.write(source, str(target))
    archive.writestr(str(root / 'README.txt'), readme.encode('utf-8'))

with ZipFile(archive_path) as archive:
    assert archive.testzip() is None, 'ZIP integrity failure'
    expected = {str(path) for path in sources} | {str(root / 'README.txt')}
    assert set(archive.namelist()) == expected, 'Archive inventory mismatch'
    assert len(archive.namelist()) == len(expected), 'Duplicate entries'
    for target, source in sources.items():
        assert archive.read(str(target)) == source.read_bytes(), f'Archive source mismatch: {target}'
    for target in expected:
        assert '..' not in PurePosixPath(target).parts
        assert not PurePosixPath(target).is_absolute()
    # Relative skill references must remain usable after extraction.
    for relative in packaged_skill_files:
        if not relative.endswith('.md'):
            continue
        source_target = root / 'skills' / skill_dir.name / relative
        body = archive.read(str(source_target)).decode('utf-8')
        for link in re.findall(r'\]\(([^)]+)\)', body):
            if '://' in link or link.startswith('#'):
                continue
            path = link.split('#', 1)[0]
            resolved = source_target.parent / path
            assert str(resolved) in expected, f'Missing linked resource: {resolved}'

print(archive_path)
print(f'Verified {len(expected)} files; {len(packaged_skill_files)} skill files; '
      f'{archive_path.stat().st_size:,} bytes.')
