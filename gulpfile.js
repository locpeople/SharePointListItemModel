const gulp = require("gulp");

const gulpBabel = require('gulp-babel');
const ts = require('gulp-typescript');
const sourcemaps = require("gulp-sourcemaps");
const mocha = require("gulp-mocha");

gulp.task("build", function () {
    var tsproject = ts.createProject("tsconfig.json");
    var tsresult = gulp.src("./SPListItemModel.ts")
        .pipe(tsproject());

    tsresult.js
        .pipe(gulpBabel({presets: ['es2015']}))
        .pipe(gulp.dest("./dist"))

    tsresult.dts.pipe(gulp.dest("./dist"));
});

gulp.task("build-test",["build"], function () {
    var tsproject = ts.createProject("tsconfig.json");
    var tsresult = gulp.src("./test/**/*.test.ts")
        .pipe(sourcemaps.init())
        .pipe(tsproject());

    tsresult.js
        .pipe(gulpBabel({presets: ['es2015']}))
        .pipe(sourcemaps.write())
        .pipe(gulp.dest("./dist/test"))

    tsresult.dts.pipe(gulp.dest("./dist/test"));
});

gulp.task("test", ["build-test"],function() {
    return gulp.src('./dist/test/**/*.test.js')
        .pipe(mocha({ui: 'bdd'}))
});