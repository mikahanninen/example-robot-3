# Release Notes for shared libraries

Heavy recommendation to use semver.org:

- Simple three digit version number signals the type of change:
- First digit: `Major` -update where some existing feature changes behaviour or is removed meaning the using ends needs to do changes.
- Second digit: `Minor` update to bring in new features and functionality that should NOT break compatibility
- Third digit:`Patch` -fix to bring in security updates

## 1.0.0

- Initial version

## 2.0.0

- Updated Python to 3.10.12
- Update pip to 23.2.1
- Updated rpaframework to 28.3.0
