// <copyright file="Flashcard.cs" company="Dustin Ellis">
// Copyright (c) Dustin Ellis. All rights reserved.
// </copyright>

namespace Flashcard_Generator
{
    /// <summary>
    /// Flashcard object.
    /// </summary>
    internal struct Flashcard
    {
        /// <summary>
        /// Gets or sets the Question.
        /// </summary>
        public string Question{ get; set; }

        /// <summary>
        /// Gets or sets the Answer.
        /// </summary>
        public string Answer { get; set; }

        /// <summary>
        /// Gets or sets the Chapter number that the material originated from.
        /// </summary>
        public int Chapter { get; set; }

        /// <summary>
        /// Gets or sets the Section number within the <see cref="Chapter"/> the material originated from.
        /// </summary>
        public int Section { get; set; }

        /// <summary>
        /// Gets or sets the Priority of the material.
        /// </summary>
        /// <remarks>
        /// The higher the number the greater the importance of the material.
        /// </remarks>
        public int Priority { get; set; }
    }
}
