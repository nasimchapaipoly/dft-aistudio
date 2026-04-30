# Firebase Security Specification

## Data Invariants
- Notices must have a valid date and text in both languages.
- People (Teachers/Staff) must have a designated type.
- Students must have unique roll numbers (enforced by application logic/document ID).
- Read access is public for notices, people, and labs.
- Read access for students is restricted to authenticated users (e.g. teachers/admins) or the student themselves (if we implement student login later, for now we'll keep it simple or restricted to a specific admin email).

## The "Dirty Dozen" Payloads (Targets for PERMISSION_DENIED)

1. **Identity Spoofing**: Attempting to create a notice with a missing required field.
2. **Identity Spoofing**: Attempting to set an `isAdmin` field if it existed on a user profile (system-only field).
3. **Identity Integrity**: Write to `notices` without being an admin.
4. **State Shortcutting**: Updating a notice's `date` to a future date while it was immutable (if we make it so).
5. **Resource Poisoning**: Document IDs > 128 chars.
6. **Value Poisoning**: `text_en` > 5000 chars.
7. **PII Leak**: Unauthorized read of `students` collection.
8. **Malicious Update**: Removing `name_en` from a `Person` document.
9. **Relational Sync**: Creating a student in a non-existent `semester` (if semesters were a collection).
10. **Shadow Update**: Adding a `hidden_field` to a Notice.
11. **Type Violation**: Setting `date` to a number.
12. **Boundary Violation**: `images` array with 100+ elements.

## Test Runner (firestore.rules.test.ts placeholder)
We will implement security rules that reject these payloads.
