enum ErrorType{
    File = "File Exists",

}

class FileError extends Error {
    type: ErrorType;

    constructor(message: string) {
        super(message); // Call the constructor of the base class `Error`
        this.name = "FileError"; // Set the error name to your custom error class name
// Set the prototype explicitly to maintain the correct prototype chain
        Object.setPrototypeOf(this, FileError.prototype);
    }
}