<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="AppPartContent.aspx.cs" Inherits="TestAppPartWeb.AppPartContent" %>

<!DOCTYPE html>

<html>
    <body>
        <div ng-app="myApp">
	        <div ng-view>
	        </div>
        </div>

    <!-- Main JavaScript function, controls the rendering
         logic based on the custom property values -->
        <script type="text/javascript" src="https://cdnjs.cloudflare.com/ajax/libs/angular.js/1.0.8/angular.min.js"></script>
    <script type="text/javascript">
        var app = angular.module('myApp', []);

        app.controller('WelcomeController', ['$scope', function ($scope) {
            $scope.welcomeText = "World!";
        }]);

        app.controller('AboutController', ['$scope', function ($scope) {
            $scope.aboutText = "This could be an about page!";
        }]);

        app.config(['$routeProvider',
		  function ($routeProvider) {
		      $routeProvider.
                when('/home', {
                    template: '<a href="#about">About</a><a href="#home">Home</a><div><p>Hello {{welcomeText}}</p></div>',
                    controller: 'WelcomeController'
                }).
                when('/about', {
                    template: '<a href="#about">About</a><a href="#home">Home</a><div><p>{{aboutText}}</p></div>',
                    controller: 'AboutController'
                }).
                otherwise({
                    redirectTo: '/home'
                });
		  }]);
	</script>

    </body>
</html>
