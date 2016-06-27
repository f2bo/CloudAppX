var fs = require('fs'),
    Q = require('q'),
    child_process = require('child_process'),
    util = require('util'),
    path = require('path'),
    unzip2 = require('unzip2'),
    os = require('os'),
    rmdir = Q.nfbind(require('rimraf')),
    execute = Q.nfbind(child_process.exec),
    spawn = child_process.spawn,
    fsStat = Q.nfbind(fs.stat),
    readdir = Q.nfbind(fs.readdir);

var defaultToolsFolder = 'appxsdk';

function getAppx(file) {
  var ctx;
  
  return Q(file.xml)
    // unzip package content
    .then(getContents)
    // generate PRI file
    .then(function (file) {
      ctx = file;
      return makePri(file);
    })
    // move PRI file into package folder
    .then(function (file) {
      var targetPath = path.resolve(file.dir, path.basename(file.out));
      return Q.nfcall(fs.rename, file.out, targetPath).thenResolve(ctx);
    })
    // generate APPX file
    .then(function (file) {
      return makeAppx(file);
    })
    // clean up package contents
    .finally(function () {
      if (ctx) {
        return deleteContents(ctx);
      }
    });
}

function getPri(file) {
  var ctx;
  
  return Q(file.xml)
    // unzip package content
    .then(getContents)
    // generate PRI file
    .then(function (file) {
      ctx = file;
      return makePri(ctx);
    })
    // clean up package contents
    .finally(function (file) {
      if (ctx) {
        return deleteContents(ctx);
      }
    });
}

// search for local installation of Windows 10 Kit in the Windows registry
function getWindowsKitPath(toolname) {
  var cmdLine = 'powershell -noprofile -noninteractive -Command "Get-ItemProperty \\"HKLM:\\SOFTWARE\\Microsoft\\Windows Kits\\Installed Roots\\" -Name KitsRoot10 | Select-Object -ExpandProperty KitsRoot10"';
  return execute(cmdLine)
    .then(function (args) {
      var toolPath = path.resolve(args[0].replace(/[\n\r]/g, ''), 'bin', os.arch(), toolname);
      return fsStat(toolPath)
                .thenResolve(toolPath);
    })
    .catch(function (err) {
      return Q.reject(new Error('Cannot find the Windows 10 SDK tools.'));
    });
}

// search for local installation of Windows 10 tools in app's subfolder
function getLocalToolsPath(toolName) {
  // test WEBSITE_SITE_NAME environment variable to determine if the service is running in Azure, which  
  // requires mapping the tool's location to its physical path using the %HOME_EXPANDED% environment variable
  var toolPath = process.env.WEBSITE_SITE_NAME ?
                  path.join(process.env.HOME_EXPANDED, 'site', 'wwwroot', defaultToolsFolder, toolName) :
                  path.join(path.dirname(require.main.filename), defaultToolsFolder, toolName);
  
  return fsStat(toolPath)
    .thenResolve(toolPath)
    .catch(function (err) {
      return Q.reject(new Error('Cannot find Windows 10 Kit Tools in the app folder (' + defaultToolsFolder + ').'));
    });
}

// reads an app manifest and returns the package identity
// see https://msdn.microsoft.com/en-us/library/windows/apps/br211441.aspx
function getPackageIdentity(manifestPath) {
  // defines a globally unique identifier for a package
  var identityElement = /<Identity\s+[^>]+\>/;

  // A string between 3 and 50 characters in length that consists of alpha-numeric, period, and dash characters
  var nameAttribute = /Name="([A-Za-z0-9\-\.]+?)"/;

  return Q.nfcall(fs.readFile, manifestPath).then(function (data) {
    var identityMatch = data.toString().match(identityElement);
    if (identityMatch) {
      var nameMatch = identityMatch[0].match(nameAttribute);
      if (nameMatch) {
        return nameMatch[1];
      }
    }
  });
}

// generates a resource index file (PRI)
function makePri(file) {
  if (os.platform() !== 'win32') {
    return Q.reject(new Error('Cannot index Windows resources in the current platform.'));
  }
  
  var toolName = 'makepri.exe';
  var priFilePath = path.join(file.out, 'resources.pri');
  return Q.nfcall(fs.unlink, priFilePath).catch(function (err) {
            // delete existing file and report any error other than not found
            if (err.code !== 'ENOENT') {
              throw err;
            }    
          })
          .then (function () {
            return getLocalToolsPath(toolName).catch(function (err) {
              return getWindowsKitPath(toolName);
            })
            .then(function (toolPath) {
              var manifestPath = path.join(file.dir, 'appxmanifest.xml');
              return getPackageIdentity(manifestPath).then(function (packageIdentity) {
                var deferred = Q.defer();
                var configPath = path.resolve(__dirname, '..', 'assets', 'priconfig.xml');
                var process = spawn(toolPath, 
                                    ['/o',  '/pr', file.dir, '/cf', configPath, '/of ', priFilePath, '/in', packageIdentity]);
                
                var stdout = '', stderr = '';
                process.stdout.on('data', function (data) { stdout += data; });
                process.stderr.on('data', function (data) { stderr += data; });
                
                process.on('close', (code) => {
                  if (code !== 0) {
                    var toolErrors = stdout.match(/error.*/g);
                    var errmsg = toolErrors ? toolErrors.map(function (item) { return item.replace(/error:*\s*/, ''); }) : 'MakePri failed.';
  	                return deferred.reject(new Error(errmsg));
                  }

                  deferred.resolve({
                    dir: file.dir,
                    out: priFilePath,
                    stdout: stdout,
                    stderr: stderr,
                    code: code
                  });
                });
                
                process.on('error', function (err) {
	                deferred.reject(err);
                });

                return deferred.promise;
              });
            });
          })
}

function makeAppx(file) {
  if (os.platform() !== 'win32') {
    return Q.reject(new Error('Cannot generate a Windows Store package in the current platform.'));
  }
  
  var toolName = 'makeappx.exe';
  return getLocalToolsPath(toolName)
          .catch(function (err) {
            return getWindowsKitPath(toolName);
          })
          .then(function (toolPath) {
            var deferred = Q.defer();
            var packagePath = path.join(file.out, file.name + '.appx');
            var process = spawn('"' + toolPath + '"', ['pack', '/o', '/d', file.dir, '/p', packagePath, '/l']);
            
            var stdout = '', stderr = '';
            process.stdout.on('data', function (data) { stdout += data; });
            process.stderr.on('data', function (data) { stderr += data; });
            
            process.on('close', (code) => {
              if (code !== 0) {
                var toolErrors = stdout.match(/error:.*/g);
                var errmsg = toolErrors ? toolErrors.map(function (item) { return item.replace(/error:\s*/, ''); }) : 'MakePri failed.';
                return deferred.reject(new Error(errmsg));
              }

              deferred.resolve({
                dir: file.dir,
                out: packagePath,
                stdout: stdout,
                stderr: stderr,
                code: code
              });
            });
            
            process.on('error', function (err) {
              var errmsg;
              var toolErrors = stdout.match(/error:.*/g);
              if (toolErrors) {
                errmsg = toolErrors.map(function (item) { return item.replace(/error:\s*/, ''); });
              }
              else {
                errmsg = err.message;
              }

              deferred.reject(errmsg ? errmsg.join('\n') : 'MakeAppX failed.');
            });

            return deferred.promise;
          });
}

function getContents(file) {
  var deferred = Q.defer();
  var outputDir = path.join('output', path.basename(file.name, '.' + file.extension));
  fs.createReadStream(file.path)
    .on('error', function (err) {
      console.log(err);
      deferred.reject(new Error('Failed to open the uploaded content archive.'));
    })
    .pipe(unzip2.Extract({ path: outputDir }))
    .on('close', function () {
      fs.unlink(file.path, function (err) {
        if (err) {
          console.log(err);
        }
      
        var name = path.basename(file.originalname, '.' + file.extension);
        deferred.resolve({
          name: name,
          dir: path.join(outputDir, name),
          out: outputDir
        });
      });
    })
    .on('error', function (err) {
      console.log(err);
      deferred.reject(new Error('Failed to unpack the uploaded content archive.'));
    });
  
  return deferred.promise;
}

function deleteContents(ctx) {
  return rmdir(ctx.dir)
          .catch(function (err) {
            console.log('Error deleting content folder: ' + err);
          })
          .then(function () {
            return readdir(ctx.out);
          })
          .then(function (files) {
            if (files.length === 0) {
              return rmdir(ctx.out)
            }
          })
          .catch(function (err) {
            console.log('Error deleting output folder: ' + err);
          });
}

module.exports = { getAppx: getAppx, getPri: getPri, makeAppx: makeAppx, makePri: makePri };
