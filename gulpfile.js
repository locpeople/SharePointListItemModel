const gulp = require("gulp");

const gulpBabel = require('gulp-babel');
const ts = require('gulp-typescript');

gulp.task("build", function () {
    var tsproject = ts.createProject("tsconfig.json");
    var tsresult = gulp.src("./*SPListItemModel.ts")
        .pipe(tsproject());

    tsresult.js
        .pipe(gulpBabel({presets: ['es2015']}))
        .pipe(gulp.dest("./dist"))

    tsresult.dts.pipe(gulp.dest("./dist"));
});