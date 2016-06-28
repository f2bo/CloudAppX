var fs = require('fs'),
    Q = require('q'),
    exec = require('child_process').exec,
    util = require('util'),
    path = require('path'),
    unzip2 = require('unzip2'),
    os = require('os'),
    rmdir = Q.nfbind(require('rimraf')),
    execute = Q.nfbind(exec),
    fsStat = Q.nfbind(fs.stat),
    readdir = Q.nfbind(fs.readdir);

var defaultToolsFolder = 'appxsdk';

function getAppx(file, runMakePri) {
  var ctx;
  return Q(file.xml)
    // unzip package content
    .then(getContents)
    // generate PRI file
    .then(function (fileInfo) {
      ctx = fileInfo;
      if (!runMakePri) return fileInfo;
      return makePri(fileInfo, true).then(function (priFile) {
        // move PRI file into package folder
        var targetPath = path.join(fileInfo.dir, path.basename(priFile.out));
        return Q.nfcall(fs.rename, priFile.out, targetPath).thenResolve(fileInfo);
      });
    })
    // generate APPX file
    .then(makeAppx)
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
    .then(function (fileInfo) {
      ctx = file;
      return makePri(fileInfo);
    })
    // clean up package contents
    .finally(function () {
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
function makePri(fileInfo, splitFiles) {
  if (os.platform() !== 'win32') {
    return Q.reject(new Error('Cannot index Windows resources in the current platform.'));
  }
  
  var toolName = 'makepri.exe';
  var priFilePath = path.join(fileInfo.dir, 'resources.pri');
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
      var manifestPath = path.join(fileInfo.dir, 'appxmanifest.xml');
      return getPackageIdentity(manifestPath).then(function (packageIdentity) {
        var deferred = Q.defer();
        var configPath = path.resolve(__dirname, '..', 'assets', splitFiles ? 'priconfig.split.xml' : 'priconfig.xml');
        var cmdLine = '"' + toolPath + '" new /o /pr ' + fileInfo.dir + ' /cf ' + configPath + ' /of ' + priFilePath + ' /in ' + packageIdentity;
        exec(cmdLine, { maxBuffer: 1024*1024 }, function (err, stdout, stderr) {             
          if (err) {
            console.log(err.message);
            return deferred.reject(err);
          }
  
          deferred.resolve({
            dir: fileInfo.dir,
            out: priFilePath,
            stdout: stdout,
            stderr: stderr
          });
        });

        return deferred.promise;
      });
    });
  })
}

function makeAppx(fileInfo) {
  if (os.platform() !== 'win32') {
    return Q.reject(new Error('Cannot generate a Windows Store package in the current platform.'));
  }
  
  var toolName = 'makeappx.exe';
  return getLocalToolsPath(toolName).catch(function (err) {
    return getWindowsKitPath(toolName);
  })
  .then(function (toolPath) {
    var packagePath = path.join(fileInfo.out, fileInfo.name + '.appx');
    var cmdLine = '"' + toolPath + '" pack /o /d ' + fileInfo.dir + ' /p ' + packagePath + ' /l';
    var deferred = Q.defer();
    exec(cmdLine, { maxBuffer: 1024*1024 }, function (err, stdout, stderr) {             
      if (err) {
        console.log(err.message);

        var errmsg;
        var toolErrors = stdout.match(/error:.*/g);
        if (toolErrors) {
          errmsg = stdout.match(/error:.*/g).map(function (item) { return item.replace(/error:\s*/, ''); });
        }
        return deferred.reject(errmsg ? errmsg.join('\n') : 'MakeAppX failed.');
      }

      deferred.resolve({
        dir: fileInfo.dir,
        out: packagePath,
        stdout: stdout,
        stderr: stderr
      });
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
